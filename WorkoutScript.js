function doGet() {
  var html = HtmlService.createHtmlOutputFromFile('WorkoutForm')
    .setTitle('💪 Workout Logger')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}

function getExerciseList() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Exercises');
    if (!sheet) {
      throw new Error('Exercises sheet not found');
    }
    
    // Read both Type (column A) and Exercise (column B) starting from row 2
    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, 2).getValues();
    
    // Format as "Type - Exercise" and filter out empty rows
    var formattedExercises = data
      .filter(row => row[0] && row[1]) // Remove rows where either column is empty
      .map(row => row[0] + ' - ' + row[1]); // Format as "Type - Exercise"
    
    return formattedExercises;
  } catch (error) {
    Logger.log('Error in getExerciseList: ' + error.toString());
    return [];
  }
}

function getHistoricalData(exercise) {
  try {
    var statsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Stats');
    var workoutSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Workouts');
    
    // Convert the search exercise to lowercase for case-insensitive comparison
    var searchExercise = exercise.toLowerCase();
    var exerciseRecords = [];
    
    // Get most recent workout data first
    var mostRecentWorkout = null;
    if (workoutSheet) {
      var workoutData = workoutSheet.getRange(2, 1, Math.max(1, workoutSheet.getLastRow()-1), 6).getValues();
      var mostRecentDate = null;
      
      workoutData.forEach(function(row) {
        // Convert the workout exercise name to lowercase for comparison
        var workoutExercise = (row[1] || '').toString().toLowerCase(); // Column B is Exercise
        
        if (workoutExercise === searchExercise && row[3] && row[4]) { // Has weight and reps data
          var rowDate = new Date(row[0]); // Column A is Date
          
          if (!mostRecentDate || rowDate > mostRecentDate) {
            mostRecentDate = rowDate;
            mostRecentWorkout = {
              weight: Number(row[3]) || 0,  // Column D is Weight
              reps: row[4] || '',           // Column E is Reps (could be comma-separated)
              rir: row[5] || '',            // Column F is RIR (could be comma-separated)
              date: Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'MM/dd/yyyy'),
              isRecentWorkout: true
            };
          }
        }
      });
    }
    
    // Get historical stats data
    if (statsSheet) {
      var statsData = statsSheet.getRange(2, 1, Math.max(1, statsSheet.getLastRow()-1), 4).getValues();
      
      statsData.forEach(function(row) {
        // Convert the stats exercise name to lowercase for comparison
        var statsExercise = (row[0] || '').toString().toLowerCase();
        
        if (statsExercise === searchExercise && row[1] && row[2]) { // Exercise matches and has weight/reps data
          exerciseRecords.push({
            weight: Number(row[1]) || 0,
            reps: Number(row[2]) || 0,
            date: row[3] ? Utilities.formatDate(new Date(row[3]), Session.getScriptTimeZone(), 'MM/dd/yyyy') : '',
            isRecentWorkout: false
          });
        }
      });
    }
    
    // Sort historical records by date (most recent first) then by weight (highest first)
    exerciseRecords.sort(function(a, b) {
      if (a.date !== b.date) {
        return new Date(b.date) - new Date(a.date);
      }
      return b.weight - a.weight;
    });
    
    // Add most recent workout at the top if found
    var finalRecords = [];
    if (mostRecentWorkout) {
      finalRecords.push(mostRecentWorkout);
    }
    finalRecords = finalRecords.concat(exerciseRecords);
    
    Logger.log("Historical data for " + exercise + " (with recent workout): " + JSON.stringify(finalRecords));
    return JSON.stringify(finalRecords);
  } catch (error) {
    Logger.log('Error in getHistoricalData: ' + error.toString());
    return JSON.stringify([]);
  }
}

function addWorkout(workout) {
  try {
    var workoutSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Workouts');
    
    if (!workoutSheet) {
      throw new Error('Workouts sheet not found');
    }
    
    var currentDate = new Date();
    var dateOnly = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'MM/dd/yyyy');
    
    // Collect all rows to insert (we'll insert them in one batch)
    var rowsToInsert = [];
    
    workout.exercises.forEach(function(ex) {
      // Split "Type - Exercise" back into separate parts
      var parts = ex.exercise.split(' - ');
      var muscle = parts[0] || '';  // e.g., "Back"
      var exercise = parts[1] || ex.exercise;  // e.g., "Cable rows"
      
      // Group sets by weight to create separate rows for different weights
      var weightGroups = {};
      
      ex.sets.forEach(function(set) {
        var weight = set.weight;
        if (!weightGroups[weight]) {
          weightGroups[weight] = {
            reps: [],
            rirs: []
          };
        }
        weightGroups[weight].reps.push(set.reps);
        weightGroups[weight].rirs.push(set.rir || ''); // Handle empty RIR values
      });
      
      // Create a separate row for each weight used
      Object.keys(weightGroups).forEach(function(weight) {
        var groupData = weightGroups[weight];
        var repsString = groupData.reps.join(',');
        var rirsString = groupData.rirs.join(',');
        
        // Prepare row data
        // Format: Date | Exercise | Muscle | Weight | Reps | RIR
        rowsToInsert.push([
          dateOnly,        // Date (MM/dd/yyyy format)
          exercise,        // Exercise name only
          muscle,          // Muscle group
          Number(weight),  // Weight (single value for this row)
          repsString,      // Reps for this weight (comma-separated)
          rirsString       // RIR ratings for this weight (comma-separated)
        ]);
      });
    });
    
    // Insert all rows at once after the header (row 2)
    if (rowsToInsert.length > 0) {
      workoutSheet.insertRows(2, rowsToInsert.length);
      var range = workoutSheet.getRange(2, 1, rowsToInsert.length, 6);
      range.setValues(rowsToInsert);
      
      // Add a black border around the entire workout session (no background color)
      // Highlighting all 6 columns: Date | Exercise | Muscle | Weight | Reps | RIR
      range.setBorder(
        true,  // top
        true,  // left  
        true,  // bottom
        true,  // right
        false, // vertical (no internal vertical lines)
        false, // horizontal (no internal horizontal lines)
        'black', 
        SpreadsheetApp.BorderStyle.SOLID
      );
    }
    
    return 'Success';
  } catch (error) {
    Logger.log('Error in addWorkout: ' + error.toString());
    throw new Error('Failed to save workout: ' + error.toString());
  }
}

