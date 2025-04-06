function buildReports(forms) {
  const allFormData = compileFormResponses(forms, 'reports');
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

  // Step 3: Find Top & Bottom Fliers (Only Include Students in Last 5 Exit Tickets)
  const studentFliers = buildStudentFliers(studentRecords);

  // Step 4: Compute Teacher Period Averages (Exclude invalid periods)
  const teacherPeriodAverages = buildTeacherAverages(teacherAverages);  

  // Step 5: Determine Top and Bottom 10% Students (Based on the Whole Year)
  const bottomAndTopStudents = buildTopAndBottomTenPercentStudents(studentRecords);

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

function buildStudentFliers (studentRecords) {
  const topFliers = [];
  const bottomFliers = [];
  
  Object.entries(studentRecords).forEach(([email, responses]) => {
    if (responses.length < 3) return; // Need at least 3 responses
  
    const lastThree = responses.slice(-3);
    const overallAvg = responses.reduce((sum, s) => sum + (s.score / s.possibleScore), 0) / responses.length;
    const lastAvg = lastThree.reduce((sum, s) => sum + (s.score / s.possibleScore), 0) / lastThree.length;
    const percentChange = (lastAvg - overallAvg) / overallAvg;
  
    if (percentChange >= 0.25) {
      topFliers.push({ 
        email, 
        studentName: responses[0].studentName,
        teacher: responses[0].teacher
      });
    } else if (percentChange <= -0.25) {
      bottomFliers.push({ 
        email, 
        studentName: responses[0].studentName,
        teacher: responses[0].teacher
      });
    };
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

function buildTopAndBottomTenPercentStudents(studentRecords) {
  const uniqueStudents = Object.entries(studentRecords).map(([email, responses]) => {
    const totalScore = responses.reduce((sum, s) => sum + s.score, 0);
    const totalPossible = responses.reduce((sum, s) => sum + s.possibleScore, 0);

    return totalPossible > 0 ? { 
        email, 
        studentName: responses[0].studentName, 
        teacher: responses[0].teacher,
        avgPercentage: totalScore / totalPossible 
    } : null;
  }).filter(Boolean); // Remove null values

  uniqueStudents.sort((a, b) => b.avgPercentage - a.avgPercentage);

  const numStudents = uniqueStudents.length;
  const topCount = Math.ceil(numStudents * 0.1);
  const bottomCount = Math.ceil(numStudents * 0.1);

  const top10StudentsList = uniqueStudents.slice(0, topCount).map(s => ({ 
    email: s.email, 
    studentName: s.studentName, 
    teacher: s.teacher
  })).sort((a, b) => a.teacher.localeCompare(b.teacher));

  const bottom10StudentsList = uniqueStudents.slice(-bottomCount).map(s => ({ 
    email: s.email, 
    studentName: s.studentName, 
    teacher: s.teacher
  })).sort((a, b) => a.teacher.localeCompare(b.teacher));

  return {
    top10Students: top10StudentsList,
    bottom10Students: bottom10StudentsList,
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

      return [`${student.studentName.trim()} --- ${student.email} --- ${student.teacher}`];
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