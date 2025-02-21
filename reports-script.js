function buildReports(forms) {
    const allFormData = compileFormResponses(forms);
    const allStudents = [];
    let teacherAverages = {};
    let formStrandPerformance = {};
    let studentRecords = {};
  
    function validatePeriod(period) {
      return /^[0-9]+$/.test(period) ? period : "";
    };
  
    // Collect student responses
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
        }
        if (!formStrandPerformance[studentObj.formStrand][studentObj.teacher]) {
          formStrandPerformance[studentObj.formStrand][studentObj.teacher] = { totalScore: 0, totalPossible: 0, count: 0 };
        }
        formStrandPerformance[studentObj.formStrand][studentObj.teacher].totalScore += studentObj.score;
        formStrandPerformance[studentObj.formStrand][studentObj.teacher].totalPossible += studentObj.possibleScore;
        formStrandPerformance[studentObj.formStrand][studentObj.teacher].count++;
  
        // Track teacher period averages (only if period is valid)
        if (period) {
          const teacherPeriodKey = `${period}-${studentObj.teacher}`;
          if (!teacherAverages[teacherPeriodKey]) {
            teacherAverages[teacherPeriodKey] = { totalScore: 0, totalPossible: 0 };
          }
          teacherAverages[teacherPeriodKey].totalScore += studentObj.score;
          teacherAverages[teacherPeriodKey].totalPossible += studentObj.possibleScore;
        };
      });
    });
  
    // 1️⃣ Find Top & Bottom Fliers
    const topFliers = [];
    const bottomFliers = [];
  
    Object.entries(studentRecords).forEach(([email, responses]) => {
      if (responses.length < 3) return; // Need at least 3 responses
  
      let lastThree = responses.slice(-3);
      let overallAvg = responses.reduce((sum, s) => sum + (s.score / s.possibleScore), 0) / responses.length;
      let lastAvg = lastThree.reduce((sum, s) => sum + (s.score / s.possibleScore), 0) / lastThree.length;
      let percentChange = (lastAvg - overallAvg) / overallAvg;
  
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
  
    // 2️⃣ Compute Teacher Period Averages (Exclude invalid periods)
    let teacherPeriodAverages = Object.keys(teacherAverages)
      .map(key => {
        let { totalScore, totalPossible } = teacherAverages[key];
        return totalPossible > 0 ? { teacherPeriod: key, avgScore: totalScore / totalPossible } : null;
      })
      .filter(Boolean); // Remove null values
  
    // 3️⃣ Determine Top and Bottom 10% Students
    let uniqueStudents = Object.entries(studentRecords).map(([email, responses]) => {
      let totalScore = responses.reduce((sum, s) => sum + s.score, 0);
      let totalPossible = responses.reduce((sum, s) => sum + s.possibleScore, 0);
      
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
  
    const top10Students = uniqueStudents.slice(0, topCount).map(s => ({ 
      email: s.email, 
      studentName: s.studentName, 
      teacher: s.teacher
    }));
  
    const bottom10Students = uniqueStudents.slice(-bottomCount).map(s => ({ 
      email: s.email, 
      studentName: s.studentName, 
      teacher: s.teacher
    }));
  
    // 4️⃣ Find Best & Worst Teachers Per Strand
    let bestWorstTeachers = {};
  
    Object.keys(formStrandPerformance).forEach(formStrand => {
      let teachers = Object.keys(formStrandPerformance[formStrand]).map(teacher => {
        let data = formStrandPerformance[formStrand][teacher];
        return data.totalPossible > 0
          ? { teacher, avgScore: data.totalScore / data.totalPossible }
          : null;
      }).filter(Boolean); // Remove null values
  
      if (teachers.length > 1) {
        teachers.sort((a, b) => b.avgScore - a.avgScore); // Highest to lowest
        bestWorstTeachers[formStrand] = {
          bestTeacher: teachers[0].teacher,
          worstTeacher: teachers[teachers.length - 1].teacher
        };
      } else if (teachers.length === 1) {
        bestWorstTeachers[formStrand] = {
          bestTeacher: teachers[0].teacher,
          worstTeacher: null
        };
      };
    });
  
    return {
      topFliers,
      bottomFliers,
      teacherPeriodAverages,
      top10Students,
      bottom10Students,
      bestWorstTeachers
    };
  };
  
  function fillReportsSheet(data) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Reports");
  
    if (!sheet) {
      sheet = ss.insertSheet("Reports");
    };
    
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
      const studentData = studentList.map(student => [
        `${student.studentName.trim()} --- ${student.email} --- ${student.teacher}`
      ]);
      if (studentData.length > 0) {
        sheet.getRange(row, col, studentData.length, 1).setValues(studentData);
      };
    };
    
    fillColumnData(4, data.bottom10Students);
    fillColumnData(5, data.top10Students);
    fillColumnData(6, data.bottomFliers);
    fillColumnData(7, data.topFliers);
    
    // Fill Best/Worst Teachers by Strand
    const strandData = [];
    Object.keys(data.bestWorstTeachers).forEach(strand => {
      const entry = data.bestWorstTeachers[strand];
      strandData.push([strand, entry.bestTeacher, entry.worstTeacher]);
    });
    if (strandData.length > 0) {
      sheet.getRange(row, 9, strandData.length, 3).setValues(strandData);
    };
  };