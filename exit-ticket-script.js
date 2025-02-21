function fillGoogleSheet(sheet, forms, allFormData) {
    let row = 3;
    let currentFormIndex = 0;
  
    buildExitTicketFrames(sheet, forms, allFormData);
  
    allFormData.forEach(formData => { // start filling out new columns and rows for each form
  
      if (currentFormIndex !== formData.formOrder) { // check if new form has started, if so reset rows
        currentFormIndex = formData.formOrder;
        row = 3;
      };
  
      let currentPeriod = "1"; // must be set to whatever the first period is
      let switchStatus = false;
      const whiteHex = "#FFFFFF";
      const greyHex = "#F0F0F0";
  
      formData["allStudentResponses"].forEach(student => {
  
        if (student.period !== currentPeriod) {
          currentPeriod = student.period;
          switchStatus = !switchStatus;
        }
  
        // teacher name cell
        sheet.getRange(row, forms[formData.formOrder].startColumn).setValue(`${!isNaN(Number(student.period)) ? student.period : ""}-${student.teacherName}`);
        sheet.getRange(row, forms[formData.formOrder].startColumn).setBackground(switchStatus === true ? whiteHex : greyHex);
  
        // student name cell
        sheet.getRange(row, forms[formData.formOrder].startColumn + 1).setValue(student.studentName);
        sheet.getRange(row, forms[formData.formOrder].startColumn + 1).setBackground(switchStatus === true ? whiteHex : greyHex);
  
        // email cell
        sheet.getRange(row, forms[formData.formOrder].startColumn + 2).setValue(student.email);
        sheet.getRange(row, forms[formData.formOrder].startColumn + 2).setBackground(switchStatus === true ? whiteHex : greyHex);
  
        // score cell
        sheet.getRange(row, forms[formData.formOrder].startColumn + 3).setValue(student.score);
        sheet.getRange(row, forms[formData.formOrder].startColumn + 3).setBackground(getScoreBackgroundColor(student.score, student.possibleScore));
  
        row++;
      });
    });
  
    const averageScores = buildTeacherAverages(allFormData);
    fillTeacherAverages(sheet, averageScores);
  };
  
  function buildExitTicketFrames(sheet, forms, formData) {
    // forms has: url, formOrder, and startColumn
    // responses has: allStudentResponses array, formTitle, and formOrder
    const darkGrayHex = "#C0C0C0";
  
    // Determine the required number of columns
    let maxColumn = 0;
    for (let i = 0; i < formData.length; i++) {
      let exitTicketData = formData[i];
      let endColumn = forms[exitTicketData.formOrder].startColumn + 4;
      if (endColumn > maxColumn) {
        maxColumn = endColumn;
      }
    }
  
    // Expand sheet columns if needed
    if (sheet.getMaxColumns() < maxColumn) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), maxColumn - sheet.getMaxColumns());
    };
  
    for (let i = 0; i < formData.length; i++) {
      const exitTicketData = formData[i];
      const startCol = forms[exitTicketData.formOrder].startColumn;
  
      sheet.getRange(1, startCol).setValue("Exit Ticket:").setFontWeight("bold");
      sheet.getRange(1, startCol + 1).setValue(exitTicketData.formTitle);
  
      const headers = [
        "Teacher Name", 
        "Student Name", 
        "Email", 
        "Score", 
        "Avg. Score"
      ];
  
      for (let j = 0; j < headers.length; j++) {
        let columnIndex = startCol + j;
        sheet.getRange(2, columnIndex).setValue(headers[j]).setFontWeight("bold");
        sheet.getRange(2, columnIndex).setBackground(darkGrayHex);
      };
    };
  };
  
  function buildTeacherAverages(allFormData) {
    const quizzes = [];
  
    allFormData.forEach((formData, index) => {
      if (formData.allStudentResponses 
        && Array.isArray(formData.allStudentResponses)
      ) {
        let totalScore = 0;
        let totalPossible = 0;
        let teacherScores = {};
  
        formData.allStudentResponses.forEach(student => {
          let { 
            teacherName, 
            score, 
            possibleScore 
          } = student;
  
          // Track total scores for overall average
          totalScore += score;
          totalPossible += possibleScore;
  
          // Initialize teacher entry if not exists
          if (!teacherScores[teacherName]) {
            teacherScores[teacherName] = { totalScore: 0, totalPossible: 0 };
          }
  
          // Accumulate scores for teacher
          teacherScores[teacherName].totalScore += score;
          teacherScores[teacherName].totalPossible += possibleScore;
        });
  
        // Overall average for this quiz
        let overallAverage = totalPossible > 0 ? ((totalScore / totalPossible) * 100).toFixed(1) : "0.0";
  
        // Averages per teacher for this quiz
        let teacherAverages = {};
        for (let teacher in teacherScores) {
          let { totalScore, totalPossible } = teacherScores[teacher];
          teacherAverages[teacher] = totalPossible > 0 ? ((totalScore / totalPossible) * 100).toFixed(1) : "0.0";
        };
  
        // Store this quiz's averages
        quizzes.push({
          quizIndex: index + 1, // Assign a quiz number (1-based index)
          overallAverage: overallAverage,
          teacherAverages: teacherAverages
        });
      }
    });
  
    return quizzes;
  };
  
  function fillTeacherAverages(sheet, averageScores) {
    const backgroundColor = "#E0E0E0";
    let startColumn = 5;
    let row = 3;
  
    for (let i = 0; i < averageScores.length; i++) {
      sheet.getRange(row, startColumn).setValue(averageScores[i].overallAverage);
      sheet.getRange(row, startColumn).setBackground(backgroundColor);
      row++;
  
      for (let teacher in averageScores[i].teacherAverages) {
        const teacherAverage = averageScores[i].teacherAverages[teacher];
        sheet.getRange(row, startColumn).setValue(`${teacher}`).setFontWeight("bold");
        sheet.getRange(row, startColumn).setBackground(backgroundColor);
        sheet.getRange(row + 1, startColumn).setValue(`${teacherAverage}`);
        sheet.getRange(row + 1, startColumn).setBackground(backgroundColor);
        row += 2;
      };
  
      row = 3;
      startColumn += 6;
    };
  };
  
  function getScoreBackgroundColor(score, possibleScore) {
    if (possibleScore === 0) return "#F4CCCC";
  
    const percentage = (score / possibleScore) * 100;
  
    if (percentage > 90) {
      return "#C9DAF8"; // light blue
    } else if (percentage > 75) {
      return "#D9EAD3"; // light green
    } else if (percentage > 40) {
      return "#FFF2CC"; // light yellow
    } else {
      return "#F4CCCC"; // light red
    };
  };