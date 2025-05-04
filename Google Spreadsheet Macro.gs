function Vanes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var ui = SpreadsheetApp.getUi(); 

  var sht = ss.getSheetByName("Vanes");
  if (!sht) {
    ui.alert("Sheet Tidak Ditemukan!");
    return;
  }

  while (true) { // Infinite loop agar terus meminta Task ID
    var sheet = ss.getSheetByName("Vanes");
    var data = sheet.getRange("A:A").getValues().flat();
    var rowIndex = data.indexOf(taskId);
    
    var respTask = ui.prompt("Vanes | Tempelkan Task ID di sini", ui.ButtonSet.OK_CANCEL);
    
    if (respTask.getSelectedButton() == ui.Button.CANCEL) {
      return; 
    }

    var taskId = respTask.getResponseText().trim().replace(/\s/g, "");

    if (taskId === "") {
      ui.alert("Task ID tidak boleh kosong!");
      continue; // Kembali meminta input Task ID
    }

    var sheets = ss.getSheets();
    var found = false;
    
    for (var i = 0; i < sheets.length; i++) {
      var sheet = sheets[i];
      var data = sheet.getRange("A:A").getValues().flat();
      
      var rowIndex = data.indexOf(taskId);
      if (rowIndex !== -1) { // Jika task ID ditemukan
        found = true;
        var answer = sheet.getRange(rowIndex + 1, 3).getValue();
        var alasan = sheet.getRange(rowIndex + 1, 4).getValue(); 
        sheet.getRange(rowIndex + 1, 5).setValue("Vanes"); 

        var respRelevanceNotes = ui.prompt(
          "Jawaban " + sheet.getName() + " adalah " + answer + " \n" +
          "Alasan " + alasan + " \n\n" +
          "Masukkan relevansi dan catatan tambahan kalau Verifier 1 salah:\n\n" +
          "Jika Verifier 1 benar klik ok aja:\n" +
          "Contoh:\n" +
          "2 Style sudah sesuai\n" +
          "Pilihan relevansi:\n" +
          "2 - Strongly Relevant (2)\n" +
          "1 - Weakly Relevant (1)\n" +
          "0 - Irrelevant (0)\n" +
          "-1 - Abandon (-1)",
          ui.ButtonSet.OK_CANCEL
        );

        if (respRelevanceNotes.getSelectedButton() == ui.Button.CANCEL) {
          ui.alert("Input dibatalkan.");
          return;
        }

        var inputText = respRelevanceNotes.getResponseText().trim();
        var inputParts = inputText.split(/\s+/); // Pisahkan berdasarkan spasi

        if (inputText == "") {
          sheet.getRange(rowIndex + 1, 6).setValue(sheet.getRange(rowIndex + 1, 3).getValue()).setBackground("green");
          continue;
        }

        if (inputParts.length < 1) {
          ui.alert("Angkanya Hanya Boleh -1 0 1 2! Masukkan relevansi minimal, contoh: 2 atau 2 Tugas ini sangat relevan.");
          continue;
        }

        var relevanceInput = inputParts[0].trim();
        var notes = inputParts.length > 1 ? inputParts.slice(1).join(" ").trim() : ""; 

        // Mapping relevansi ke format baru
        var relevanceMapping = {
          "2": "Strongly Relevant (2)",
          "1": "Weakly Relevant (1)",
          "0": "Irrelevant (0)",
          "-1": "Abandon (-1)"
        };

        var relevance = relevanceMapping[relevanceInput]; 

        sheet.getRange(rowIndex + 1, 6).setValue(relevance);
        sheet.getRange(rowIndex + 1, 7).setValue(notes);

        break; // Keluar dari loop pencarian, lanjut ke Task ID berikutnya
      }
    }

    if (!found) {
      // Jika Task ID tidak ditemukan, minta input relevansi & catatan dalam 1 prompt
      var respRelevanceNotes = ui.prompt(
        "Masukkan relevansi dan catatan tambahan (pisahkan dengan spasi):\n\n" +
        "Contoh:\n" +
        "2 Style sudah sesuai\n" +
        "Atau jika tanpa catatan:\n" +
        "2\n\n" +
        "Pilihan relevansi:\n" +
        "2 - Strongly Relevant (2)\n" +
        "1 - Weakly Relevant (1)\n" +
        "0 - Irrelevant (0)\n" +
        "-1 - Abandon (-1)",
        ui.ButtonSet.OK_CANCEL
      );

      if (respRelevanceNotes.getSelectedButton() == ui.Button.CANCEL) {
        ui.alert("Input dibatalkan.");
        return;
      }

      var inputText = respRelevanceNotes.getResponseText().trim();
      var inputParts = inputText.split(/\s+/); // Pisahkan berdasarkan spasi

      if (inputParts.length < 1) {
        ui.alert("Format salah! Masukkan relevansi minimal, contoh: 2 atau 2 Tugas ini sangat relevan.");
        continue;
      }

      var relevanceInput = inputParts[0].trim();
      var notes = inputParts.length > 1 ? inputParts.slice(1).join(" ").trim() : ""; 

      // Mapping relevansi ke format baru
      var relevanceMapping = {
        "2": "Strongly Relevant (2)",
        "1": "Weakly Relevant (1)",
        "0": "Irrelevant (0)",
        "-1": "Abandon (-1)"
      };

      if (!relevanceMapping.hasOwnProperty(relevanceInput)) {
        ui.alert("Pilihan relevansi tidak valid! Masukkan 2, 1, 0, atau -1.");
        continue;
      }

      var relevance = relevanceMapping[relevanceInput]; 

      var lr = sht.getLastRow() + 1;
      sht.getRange(lr, 1).setValue(taskId);
      sht.getRange(lr, 3).setValue(relevance);
      sht.getRange(lr, 4).setValue(notes);
    }
  }
}
