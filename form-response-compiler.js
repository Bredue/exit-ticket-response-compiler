function compileFormResponses(forms) {
    const formsResponses = [];
  
    forms.forEach(function(formInfo) {
      const form = FormApp.openByUrl(formInfo.url);
      const formTitle = form.getTitle();
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
  
        const studentResponse = new StudentResponse(
          formInfo.strand,
          teacherName, 
          studentName, 
          email, 
          classPeriod, 
          score, 
          possibleScore
        );
  
        allStudentResponses.push(studentResponse);
        
      });
  
      allStudentResponses.sort((a, b) => {
        if (a.teacherName < b.teacherName) return -1;
        if (a.teacherName > b.teacherName) return 1;
  
        if (a.period < b.period) return -1;
        if (a.period > b.period) return 1;
  
        return b.score - a.score;
      });
  
      const responseObject = {
        allStudentResponses: allStudentResponses,
        formOrder: formInfo.order,
        formTitle: formTitle,
      };
  
      formsResponses.push(responseObject);
    });
  
    return formsResponses;
  };
  
  class StudentResponse {
    constructor(
      strand,
      teacherName, 
      studentName, 
      email, 
      classPeriod, 
      score, 
      possibleScore
    ) {
      this.strand = strand;
      this.teacherName = teacherName;
      this.studentName = studentName;
      this.email = email.split("@")[0];
      this.period = classPeriod;
      this.score = score;
      this.possibleScore = possibleScore;
    };
  };