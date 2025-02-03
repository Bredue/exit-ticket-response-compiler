function compileFormResponses(forms) {
    const formsResponses = [];
  
    forms.forEach(function(formInfo) {
      const form = FormApp.openByUrl(formInfo.url);
      const responses = form.getResponses();
      const allStudentResponses = [];
      
      responses.forEach(function(response) {
        const itemResponses = response.getItemResponses();
        const teacherName = itemResponses[0].getResponse();
        const studentName = itemResponses[1].getResponse();
        const classPeriod = itemResponses[2].getResponse();
        const email = response.getRespondentEmail();
        
        let possibleScore = 0;
        let score = 0;
  
        const gradableResponses = response.getGradableItemResponses();
        gradableResponses.forEach((response) => {
          score += response.getScore();
          possibleScore += 1;
        });
  
        const studentResponse = new StudentResponse(teacherName, studentName, email, classPeriod, score, possibleScore);
        allStudentResponses.push(studentResponse);
        
      });
  
      allStudentResponses.sort((a, b) => {
        // sort by teacher name
        if (a.teacherName < b.teacherName) return -1;
        if (a.teacherName > b.teacherName) return 1;
  
        // 2nd sort by period
        if (a.period < b.period) return -1;
        if (a.period > b.period) return 1;
  
        // 3rd sort by score
        return b.score - a.score;
      });
  
      const responseObject = {
        formOrder: formInfo.order,
        allStudentResponses: allStudentResponses,
      };
  
      formsResponses.push(responseObject);
    });
  
    return formsResponses;
  }
  
  function fillGoogleSheet(sheet, forms, responses) {
    let row = 3;
    let currentFormIndex = 0;
  
    responses.forEach(response => {
  
      // variables for modifying background color on period change of cell
      let currentPeriod = "1"; // must be set to whatever the first period is
      let switchStatus = false;
      const whiteHex = "#FFFFFF";
      const greyHex = "#F0F0F0";
  
      // check if new form has started, if so reset rows
      if (currentFormIndex !== response.formOrder) {
        currentFormIndex = response.formOrder;
        row = 3;
      };
  
      response["allStudentResponses"].forEach(student => {
  
        if (student.period !== currentPeriod) {
          currentPeriod = student.period;
          switchStatus = !switchStatus;
        }
  
        // teacher name cell
        sheet.getRange(row, forms[response.formOrder].startColumn).setValue(`${student.period}-${student.teacherName}`);
        sheet.getRange(row, forms[response.formOrder].startColumn).setBackground(switchStatus === true ? whiteHex : greyHex);
  
        // student name cell
        sheet.getRange(row, forms[response.formOrder].startColumn + 1).setValue(student.studentName);
        sheet.getRange(row, forms[response.formOrder].startColumn + 1).setBackground(switchStatus === true ? whiteHex : greyHex);
  
        // email cell
        sheet.getRange(row, forms[response.formOrder].startColumn + 2).setValue(student.email);
        sheet.getRange(row, forms[response.formOrder].startColumn + 2).setBackground(switchStatus === true ? whiteHex : greyHex);
  
        // score cell
        sheet.getRange(row, forms[response.formOrder].startColumn + 3).setValue(student.score);
        sheet.getRange(row, forms[response.formOrder].startColumn + 3).setBackground(getScoreBackgroundColor(student.score, student.possibleScore));
  
        row++;
      });
    });
  }
  
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
    }
  };
  
  class StudentResponse {
    constructor(teacherName, studentName, email, classPeriod, score, possibleScore) {
      this.teacherName = teacherName;
      this.studentName = studentName;
      this.email = email.split("@")[0];
      this.period = classPeriod;
      this.score = score;
      this.possibleScore = possibleScore;
    }
  }