function sendStudentReports(forms) {
    const allFormData = compileFormResponses(forms);
    const studentRecords = {};
    const teacherAverages = {};
    const formAverages = {}; // Track exit ticket scores by teacher and grade level
  
    // Compile student records, teacher averages, and form-level averages
    allFormData.forEach(form => {
      const { formTitle, formOrder, allStudentResponses } = form;
  
      if (!formAverages[formOrder]) {
        formAverages[formOrder] = { totalScore: 0, totalPossible: 0, teachers: {} };
      }
  
      allStudentResponses.forEach(response => {
        const period = validatePeriod(response.period);
        const email = response.email;
  
        const studentEntry = {
          formStrand: response.strand,
          formTitle,
          formOrder,
          score: response.score,
          possibleScore: response.possibleScore,
          studentName: response.studentName,
          teacher: response.teacherName,
          period,
        };
  
        if (!studentRecords[email]) studentRecords[email] = [];
        studentRecords[email].push(studentEntry);
  
        // Build teacher averages
        if (period) {
          const key = `${period}-${response.teacherName}`;
          if (!teacherAverages[key]) {
            teacherAverages[key] = { totalScore: 0, totalPossible: 0 };
          }
          teacherAverages[key].totalScore += response.score;
          teacherAverages[key].totalPossible += response.possibleScore;
        }
  
        // Build form-level averages (grade and per teacher)
        formAverages[formOrder].totalScore += response.score;
        formAverages[formOrder].totalPossible += response.possibleScore;
  
        if (!formAverages[formOrder].teachers[response.teacherName]) {
          formAverages[formOrder].teachers[response.teacherName] = { totalScore: 0, totalPossible: 0 };
        }
        formAverages[formOrder].teachers[response.teacherName].totalScore += response.score;
        formAverages[formOrder].teachers[response.teacherName].totalPossible += response.possibleScore;
      });
    });
  
    const students = [];
  
    Object.entries(studentRecords).forEach(([email, responses]) => {
      const formMetaLookup = {};
      allFormData.forEach(fd => {
        formMetaLookup[fd.formOrder] = {
          formTitle: fd.formTitle,
          formStrand: fd.formStrand
        };
      });
  
      const completed = responses.map(r => {
        const order = r.formOrder;
        const meta = formMetaLookup[order] || {};
  
        const studentAvg = r.possibleScore > 0 ? (r.score / r.possibleScore) * 100 : 0;
  
        const teacherData = formAverages[order]?.teachers?.[r.teacher] || { totalScore: 0, totalPossible: 0 };
        const classAvg = teacherData.totalPossible > 0 ? (teacherData.totalScore / teacherData.totalPossible) * 100 : 0;
  
        const overallData = formAverages[order];
        const gradeLevelAvg = overallData.totalPossible > 0 ? (overallData.totalScore / overallData.totalPossible) * 100 : 0;
  
        return {
          formOrder: order,
          formTitle: meta.formTitle,
          formStrand: meta.formStrand,
          studentAvg,
          classAvg,
          gradeLevelAvg
        };
      });
  
      const completedFormOrders = responses.map(r => r.formOrder);
      const missingFormOrders = Object.keys(formMetaLookup)
        .map(Number)
        .filter(order => !completedFormOrders.includes(order));
  
      const missing = missingFormOrders.map(order => {
        const meta = formMetaLookup[order] || {};
        return {
          formOrder: order,
          formTitle: meta.formTitle,
          formStrand: meta.formStrand
        };
      });
  
      const total = responses.reduce((acc, r) => {
        acc.score += r.score;
        acc.possible += r.possibleScore;
        return acc;
      }, { score: 0, possible: 0 });
      const average = total.possible > 0 ? (total.score / total.possible) * 100 : 0;
  
      const freq = {};
      const nameCounts = {};
      responses.forEach(r => {
        const key = `${r.period}-${r.teacher}`;
        if (!freq[key]) freq[key] = 0;
        freq[key]++;
        const name = r.studentName;
        if (!nameCounts[name]) nameCounts[name] = 0;
        nameCounts[name]++;
      });
  
      const mostCommonName = Object.entries(nameCounts).sort((a, b) => b[1] - a[1])[0];
      const mostFrequent = Object.entries(freq).sort((a, b) => b[1] - a[1])[0];
      let [period, teacher] = mostFrequent ? mostFrequent[0].split('-') : [undefined, undefined];
  
      const classKey = `${period}-${teacher}`;
      const classData = teacherAverages[classKey];
      const classAvg = classData ? (classData.totalScore / classData.totalPossible) * 100 : 0;
  
      students.push({
        mostCommonName,
        email,
        completed,
        missing,
        average,
        period,
        teacher,
        classAverage: classAvg
      });
    });
  
    sendStudentEmails(students);
  }
  
  function validatePeriod(period) {
    return /^[0-9]+$/.test(period) ? period : "";
  }
  
  function sendStudentEmails(students) {
    students.forEach(student => {
      const completedList = student.completed.map(f =>
      `<li>
        <strong>${f.formStrand || "Untitled"}:</strong> ${f.formTitle || "Untitled"} (Exit Ticket ${Number(f.formOrder ?? 0) + 1})
        <div style="margin-left: 20px; line-height: 1.5;">
          ➤ <strong>Your Score:</strong> ${f.studentAvg.toFixed(2)}%<br>
          ➤ <strong>Class Avg:</strong> ${f.classAvg.toFixed(2)}%<br>
          ➤ <strong>Grade Level Avg:</strong> ${f.gradeLevelAvg.toFixed(2)}%
        </div>
      </li>`).join("");
  
      const missingList = student.missing.map(f =>
        `<li>${f.formStrand || "Untitled"}: ${f.formTitle || "Untitled"} (Exit Ticket ${f.formOrder})</li>`).join("");
  
      const html = `
        <div style="font-family: Arial, sans-serif; border: 1px solid #ccc; border-radius: 10px; padding: 16px;">
          <h2>Exit Ticket Report</h2>
          <p><strong>Name:</strong> ${student.mostCommonName[0]}</p>
          <p><strong>Email:</strong> ${student.email}</p>
          <p><strong>Your Exit Ticket Average:</strong> ${student.average.toFixed(2)}%</p>
          <p><strong>Class Average:</strong> ${student.classAverage.toFixed(2)}% (with ${student.teacher || 'N/A'}, Period ${student.period || 'N/A'})</p>
          <h3>✅ Completed Exit Tickets</h3>
          <ul>${completedList || "<li>None</li>"}</ul>
          <h3>❌ Missing Exit Tickets</h3>
          <ul>${missingList || "<li>None</li>"}</ul>
        </div>
      `;
  
      MailApp.sendEmail({
        to: student.email,
        subject: `${student.mostCommonName[0]}'s Exit Ticket Report - ${new Date().toLocaleDateString("en-US")} - ${student.teacher || "No Teacher"}`,
        htmlBody: html,
      });
    });
  };