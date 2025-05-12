function buildReports(forms) {
  const allFormData = compileFormResponses(forms);
  const allStudents = [];
  const teacherAverages = {};
  const formStrandPerformance = {};
  const studentRecords = {};

  function validatePeriod(period) {
    return /^[0-9]+$/.test(period) ? period : "";
  };

  // Step 1: Track students who appeared in the last 5 exit tickets
  const recentStudentAppearances = new Set();

  // Identify students who submitted responses in the last 5 exit tickets
  allFormData.slice(-5).forEach(formData => {
    formData.allStudentResponses.forEach(response => {
        recentStudentAppearances.add(response.email);
    });
  });

  // Step 2: Process All Responses
  allFormData.forEach((formData) => {
    formData.allStudentResponses.forEach((response) => {
        const period = validatePeriod(response.period);

        const studentObj = {
            email: response.email,
            studentName: response.studentName,
            teacher: response.teacherName,
            period: period,
            score: response.score,
            possibleScore: response.possibleScore,
            formStrand: response.strand,
            formTitle: formData.formTitle,
            formOrder: formData.formOrder,
        };

        allStudents.push(studentObj);

        // Track responses by student email
        if (!studentRecords[studentObj.email]) {
            studentRecords[studentObj.email] = [];
        }
        studentRecords[studentObj.email].push(studentObj);

        // Track formStrand performance per teacher
        if (!formStrandPerformance[studentObj.formStrand]) {
            formStrandPerformance[studentObj.formStrand] = {};
        };

        if (!formStrandPerformance[studentObj.formStrand][studentObj.teacher]) {
            formStrandPerformance[studentObj.formStrand][studentObj.teacher] = { totalScore: 0, totalPossible: 0, count: 0 };
        };

        formStrandPerformance[studentObj.formStrand][studentObj.teacher].totalScore += studentObj.score;
        formStrandPerformance[studentObj.formStrand][studentObj.teacher].totalPossible += studentObj.possibleScore;
        formStrandPerformance[studentObj.formStrand][studentObj.teacher].count++;

        // Track teacher period averages (only if period is valid)
        if (period) {
            const teacherPeriodKey = `${period}-${studentObj.teacher}`;

            if (!teacherAverages[teacherPeriodKey]) {
                teacherAverages[teacherPeriodKey] = { totalScore: 0, totalPossible: 0 };
            };

            teacherAverages[teacherPeriodKey].totalScore += studentObj.score;
            teacherAverages[teacherPeriodKey].totalPossible += studentObj.possibleScore;
        };
    });
  });

  // Helper function to get total number of forms
  const totalForms = getTotalNumberOfForms(studentRecords);

  // Step 3: Find Top & Bottom Fliers (Only Include Students in Last 5 Exit Tickets)
  const studentFliers = buildStudentFliers(totalForms, studentRecords);

  // Step 4: Compute Teacher Period Averages (Exclude invalid periods)
  const teacherPeriodAverages = buildTeacherAverages(teacherAverages);  

  // Step 5: Determine Top and Bottom 10% Students (Based on the Whole Year)
  const bottomAndTopStudents = buildTopAndBottomTenPercentStudents(totalForms, studentRecords);

  // Step 6: Find Best & Worst Teachers Per Strand
  const topBottomTeachers = buildTopAndBottomTeachers(formStrandPerformance);

  return {
    topFliers: studentFliers.topFliers,
    bottomFliers: studentFliers.bottomFliers,
    teacherPeriodAverages,
    top10Students: bottomAndTopStudents.top10Students,
    bottom10Students: bottomAndTopStudents.bottom10Students,
    topBottomTeachers,
  };
};

function getTotalNumberOfForms(studentRecords) {
  // Step 1: Determine max formOrder to get total number of forms
  let maxFormOrder = -1;
  Object.values(studentRecords).forEach(responses => {
    responses.forEach(r => {
      if (typeof r.formOrder === 'number' && r.formOrder > maxFormOrder) {
        maxFormOrder = r.formOrder;
      }
    });
  });
  const totalForms = maxFormOrder + 1;
  return totalForms;
};

function buildStudentFliers(totalForms, studentRecords) {
  const topFliers = [];
  const bottomFliers = [];

  // Evaluate each student
  Object.entries(studentRecords).forEach(([email, responses]) => {
    if (responses.length < 3) return; // Need at least 3 responses

    // Skip if they haven’t completed at least 50% of all available forms
    if (responses.length / totalForms < 0.5) return;

    // Sort responses by formOrder in case they're out of order
    const sortedResponses = [...responses].sort((a, b) => a.formOrder - b.formOrder);

    const lastThree = sortedResponses.slice(-3);
    const overallAvg = sortedResponses.reduce((sum, s) => sum + (s.score / s.possibleScore), 0) / sortedResponses.length;
    const lastAvg = lastThree.reduce((sum, s) => sum + (s.score / s.possibleScore), 0) / lastThree.length;
    const percentChange = (lastAvg - overallAvg) / overallAvg;

    const flierData = {
      email,
      studentName: sortedResponses[0].studentName?.trim(),
      teacher: sortedResponses[sortedResponses.length - 1].teacher,
    };

    if (percentChange >= 0.25) {
      topFliers.push(flierData);
    } else if (percentChange <= -0.25) {
      bottomFliers.push(flierData);
    }
  });

  return {
    topFliers: topFliers.sort((a, b) => a.teacher.localeCompare(b.teacher)),
    bottomFliers: bottomFliers.sort((a, b) => a.teacher.localeCompare(b.teacher)),
  };
};

function buildTeacherAverages (teacherAverages) {
  const teacherPeriodAverages = Object.keys(teacherAverages)
    .map(key => {
        const { totalScore, totalPossible } = teacherAverages[key];
        return totalPossible > 0 ? { teacherPeriod: key, avgScore: totalScore / totalPossible } : null;
    })
    .filter(Boolean); // Remove null values
  
  return teacherPeriodAverages;
}

function buildTopAndBottomTenPercentStudents(totalForms, studentRecords) {
  const qualifiedStudents = Object.entries(studentRecords).map(([email, responses]) => {
    if (responses.length / totalForms < 0.5) return null;

    const totalScore = responses.reduce((sum, s) => sum + s.score, 0);
    const totalPossible = responses.reduce((sum, s) => sum + s.possibleScore, 0);

    return totalPossible > 0 ? {
      email,
      studentName: responses[0].studentName?.trim(),
      teacher: responses[responses.length - 1].teacher,
      avgPercentage: totalScore / totalPossible
    } : null;
  }).filter(Boolean);

  // Step 3: Sort students by average percentage (high to low)
  qualifiedStudents.sort((a, b) => b.avgPercentage - a.avgPercentage);

  // Step 4: Get 10% top and bottom performers
  const numStudents = qualifiedStudents.length;
  const topCount = Math.ceil(numStudents * 0.1);
  const bottomCount = Math.ceil(numStudents * 0.1);

  const top10StudentsList = qualifiedStudents.slice(0, topCount).map(s => ({
    email: s.email,
    studentName: s.studentName,
    teacher: s.teacher,
  })).sort((a, b) => a.teacher.localeCompare(b.teacher));

  const bottom10StudentsList = qualifiedStudents.slice(-bottomCount).map(s => ({
    email: s.email,
    studentName: s.studentName,
    teacher: s.teacher,
  })).sort((a, b) => a.teacher.localeCompare(b.teacher));

  return {
    top10Students: top10StudentsList,
    bottom10Students: bottom10StudentsList
  };
};

function buildTopAndBottomTeachers(formStrandPerformance) {
  const topAndBottomTeachers = {};

  Object.keys(formStrandPerformance).forEach(formStrand => {
    const teachers = Object.keys(formStrandPerformance[formStrand]).map(teacher => {
        const data = formStrandPerformance[formStrand][teacher];
        return data.totalPossible > 0
            ? { teacher, avgScore: data.totalScore / data.totalPossible }
            : null;
    }).filter(Boolean); // Remove null values

    if (teachers.length > 1) {
        teachers.sort((a, b) => b.avgScore - a.avgScore); // Highest to lowest
        topAndBottomTeachers[formStrand] = {
            bestTeacher: teachers[0].teacher,
            worstTeacher: teachers[teachers.length - 1].teacher
        };
    } else if (teachers.length === 1) {
        topAndBottomTeachers[formStrand] = {
            bestTeacher: teachers[0].teacher,
            worstTeacher: null
        };
    };
  });

  return topAndBottomTeachers;
};

function fillReportsSheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Reports");

  if (!sheet) {
    sheet = ss.insertSheet("Reports");
  };

  clearColumns(sheet);
  
  const row = 3;

  // fill the teacher/class averages
  if (data.teacherPeriodAverages.length > 0) {
    const whiteHex = "#FFFFFF";
    const greyHex = "#F0F0F0";
    
    // Sort by teacher name, then by period number
    data.teacherPeriodAverages.sort((a, b) => {
        const [periodA, teacherA] = a.teacherPeriod.split("-");
        const [periodB, teacherB] = b.teacherPeriod.split("-");

        if (teacherA !== teacherB) {
            return teacherA.localeCompare(teacherB);
        }
        return parseInt(periodA) - parseInt(periodB);
    });

    const teacherPeriodData = [];
    const backgroundColors = [];
    
    let lastTeacher = "";
    let currentColor = whiteHex;

    data.teacherPeriodAverages.forEach(entry => {
        const [period, teacher] = entry.teacherPeriod.split("-");
        
        // Change color when teacher changes
        if (teacher !== lastTeacher) {
            currentColor = currentColor === whiteHex ? greyHex : whiteHex;
            lastTeacher = teacher;
        };
        
        teacherPeriodData.push([
            entry.teacherPeriod, 
            parseFloat((entry.avgScore * 100).toFixed(1))
        ]);

        backgroundColors.push([currentColor, currentColor]);
    });

    // Write values and apply background colors
    const range = sheet.getRange(row, 1, teacherPeriodData.length, 2);
    range.setValues(teacherPeriodData);
    range.setBackgrounds(backgroundColors);
  };
  
  // Fill Student Data
  function fillColumnData(col, studentList) {
    const whiteHex = "#FFFFFF";
    const greyHex = "#F0F0F0";
    let currentTeacher = '';
    let currentColor = whiteHex;  // Start with white
    
    const studentData = studentList.map((student, index) => {
      if (student.teacher !== currentTeacher) {
          // Switch colors when the teacher changes
          currentColor = (currentColor === whiteHex) ? greyHex : whiteHex;
          currentTeacher = student.teacher;
      };

      return [`${student.studentName.trim()} --- ${student.email.split("@")[0]} --- ${student.teacher}`];
    });

    if (studentData.length > 0) {
      const range = sheet.getRange(3, col, studentData.length, 1);
      range.setValues(studentData);

      // Apply alternating background colors based on teacher changes
      const backgroundColors = studentList.map((student, index) => {
        if (student.teacher !== currentTeacher) {
            // Switch colors when the teacher changes
            currentColor = (currentColor === whiteHex) ? greyHex : whiteHex;
            currentTeacher = student.teacher;
        };

        return [currentColor];
      });

      range.setBackgrounds(backgroundColors);
    };
  };
  
  fillColumnData(4, data.bottom10Students);
  fillColumnData(5, data.top10Students);
  fillColumnData(6, data.bottomFliers);
  fillColumnData(7, data.topFliers);
  
  // Fill Best/Worst Teachers by Strand
  const strandData = [];

  Object.keys(data.topBottomTeachers).forEach(strand => {
    const entry = data.topBottomTeachers[strand];
    strandData.push([strand, entry.bestTeacher, entry.worstTeacher]);
  });

  if (strandData.length > 0) {
    sheet.getRange(row, 9, strandData.length, 3).setValues(strandData);
  };
};

function clearColumns(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;

  const range = sheet.getRange(3, 4, lastRow - 2, 4);
  range.clear();
};