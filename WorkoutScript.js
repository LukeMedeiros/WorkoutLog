function doGet() {
  var html = HtmlService.createHtmlOutputFromFile('WorkoutForm')
    .setTitle('💪 Workout Logger')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}

// ===== EXERCISE LIST MANAGEMENT =====
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

// ===== HISTORICAL DATA =====
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

// ===== PAST WORKOUTS FUNCTIONALITY =====
function getPastWeekWorkouts() {
  try {
    var workoutSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Workouts');
    if (!workoutSheet) {
      throw new Error('Workouts sheet not found');
    }
    
    // Get all workout data first
    var data = workoutSheet.getRange(2, 1, Math.max(1, workoutSheet.getLastRow()-1), 6).getValues();
    
    // Group all workouts by date
    var allWorkoutsByDate = {};
    
    data.forEach(function(row) {
      var workoutDate = new Date(row[0]); // Column A is Date
      var dateKey = Utilities.formatDate(workoutDate, Session.getScriptTimeZone(), 'MM/dd/yyyy');
      
      if (!allWorkoutsByDate[dateKey]) {
        allWorkoutsByDate[dateKey] = {
          date: dateKey,
          dayName: Utilities.formatDate(workoutDate, Session.getScriptTimeZone(), 'EEEE'),
          actualDate: workoutDate,
          exercises: [],
          muscleGroups: new Set()
        };
      }
      
      // Parse sets information
      var repsArray = (row[4] || '').toString().split(','); // Column E is Reps
      var rirArray = (row[5] || '').toString().split(',');  // Column F is RIR
      var weight = row[3] || 0; // Column D is Weight
      var muscle = row[2] || 'Other'; // Column C is Muscle
      
      var setsInfo = repsArray.length + ' sets';
      if (repsArray.length > 0 && repsArray[0]) {
        setsInfo += ' (' + repsArray.join(', ') + ' reps @ ' + weight + ' lbs)';
      }
      
      allWorkoutsByDate[dateKey].exercises.push({
        exercise: row[1] || '', // Column B is Exercise
        muscle: muscle,
        weight: weight,
        reps: row[4] || '',     // Column E is Reps
        rir: row[5] || '',      // Column F is RIR
        setsInfo: setsInfo
      });
      
      allWorkoutsByDate[dateKey].muscleGroups.add(muscle);
    });
    
    // Convert to array and sort by date (most recent first)
    var allWorkoutsArray = Object.values(allWorkoutsByDate);
    allWorkoutsArray.sort(function(a, b) {
      return b.actualDate - a.actualDate; // Most recent first
    });
    
    // Try to get workouts from previous week first
    var today = new Date();
    var currentDay = today.getDay(); // 0 = Sunday, 1 = Monday, etc.
    var daysToLastMonday = currentDay === 0 ? 7 : currentDay; // If Sunday, go back 7 days to last Monday
    var lastMonday = new Date(today);
    lastMonday.setDate(today.getDate() - daysToLastMonday - 7); // Go back to previous week's Monday
    lastMonday.setHours(0, 0, 0, 0);
    
    var lastSunday = new Date(lastMonday);
    lastSunday.setDate(lastMonday.getDate() + 6); // Add 6 days to get Sunday
    lastSunday.setHours(23, 59, 59, 999);
    
    // Filter workouts from previous week
    var previousWeekWorkouts = allWorkoutsArray.filter(function(workout) {
      return workout.actualDate >= lastMonday && workout.actualDate <= lastSunday;
    });
    
    var finalWorkouts = [];
    
    if (previousWeekWorkouts.length >= 5) {
      // Use previous week workouts if we have at least 5
      finalWorkouts = previousWeekWorkouts.slice(0, 5);
      Logger.log('Found ' + previousWeekWorkouts.length + ' workouts from previous week, using first 5');
    } else {
      // If not enough from previous week, get the 5 most recent workouts
      finalWorkouts = allWorkoutsArray.slice(0, 5);
      Logger.log('Only found ' + previousWeekWorkouts.length + ' workouts from previous week, using 5 most recent workouts');
    }
    
    // Format the workout names with date + muscle groups
    finalWorkouts.forEach(function(workout) {
      var muscleGroupsArray = Array.from(workout.muscleGroups).sort();
      var muscleGroupsText = muscleGroupsArray.join(', ');
      
      // Create a more descriptive name
      workout.workoutName = workout.date + ' - ' + muscleGroupsText;
      workout.dayName = Utilities.formatDate(workout.actualDate, Session.getScriptTimeZone(), 'EEEE');
    });
    
    Logger.log('Returning ' + finalWorkouts.length + ' workouts');
    return JSON.stringify(finalWorkouts);
  } catch (error) {
    Logger.log('Error in getPastWeekWorkouts: ' + error.toString());
    throw new Error('Failed to load past workouts: ' + error.toString());
  }
}

function getWorkoutByDate(dateString) {
  try {
    var workoutSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Workouts');
    if (!workoutSheet) {
      throw new Error('Workouts sheet not found');
    }
    
    var targetDate = new Date(dateString);
    var data = workoutSheet.getRange(2, 1, Math.max(1, workoutSheet.getLastRow()-1), 6).getValues();
    
    var workoutData = [];
    
    data.forEach(function(row) {
      var workoutDate = new Date(row[0]); // Column A is Date
      var workoutDateString = Utilities.formatDate(workoutDate, Session.getScriptTimeZone(), 'MM/dd/yyyy');
      
      if (workoutDateString === dateString) {
        // Parse the reps and RIR data
        var repsArray = (row[4] || '').toString().split(','); // Column E is Reps
        var rirArray = (row[5] || '').toString().split(',');  // Column F is RIR
        var weight = row[3] || 0; // Column D is Weight
        
        // Create sets array
        var sets = [];
        for (var i = 0; i < repsArray.length; i++) {
          sets.push({
            reps: repsArray[i] ? repsArray[i].trim() : '',
            weight: weight,
            rir: rirArray[i] ? rirArray[i].trim() : ''
          });
        }
        
        workoutData.push({
          exercise: row[1] || '', // Column B is Exercise
          muscle: row[2] || '',   // Column C is Muscle
          sets: sets
        });
      }
    });
    
    Logger.log('Found ' + workoutData.length + ' exercises for date ' + dateString);
    return JSON.stringify(workoutData);
  } catch (error) {
    Logger.log('Error in getWorkoutByDate: ' + error.toString());
    throw new Error('Failed to load workout data: ' + error.toString());
  }
}

// ===== STATS FUNCTIONALITY =====
function getWorkoutStats() {
  try {
    var workoutSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Workouts');
    if (!workoutSheet) {
      throw new Error('Workouts sheet not found');
    }
    
    // Calculate the current week and past 3 complete calendar weeks (4 weeks total)
    var today = new Date();
    var currentDay = today.getDay(); // 0 = Sunday, 1 = Monday, etc.
    var daysToThisMonday = currentDay === 0 ? 6 : currentDay - 1; // Days back to this week's Monday
    
    var thisMonday = new Date(today);
    thisMonday.setDate(today.getDate() - daysToThisMonday);
    thisMonday.setHours(0, 0, 0, 0);
    
    Logger.log('Calculating stats from current week starting ' + thisMonday);
    
    // Get all workout data
    var data = workoutSheet.getRange(2, 1, Math.max(1, workoutSheet.getLastRow()-1), 6).getValues();
    
    // Initialize data structures
    var weeks = [];
    var muscleGroups = {};
    
    // Create week objects in descending order (most recent first)
    for (var i = 0; i < 4; i++) {
      var weekStart = new Date(thisMonday);
      weekStart.setDate(thisMonday.getDate() - (i * 7)); // Go backwards from current week
      var weekEnd = new Date(weekStart);
      weekEnd.setDate(weekStart.getDate() + 6);
      weekEnd.setHours(23, 59, 59, 999);
      
      var weekKey = 'week' + i; // Use index directly for proper ordering
      var weekLabel;
      
      if (i === 0) {
        weekLabel = 'This Week (' + Utilities.formatDate(weekStart, Session.getScriptTimeZone(), 'MM/dd') + 
                   ' - ' + Utilities.formatDate(weekEnd, Session.getScriptTimeZone(), 'MM/dd') + ')';
      } else {
        weekLabel = Utilities.formatDate(weekStart, Session.getScriptTimeZone(), 'MM/dd') + 
                   ' - ' + Utilities.formatDate(weekEnd, Session.getScriptTimeZone(), 'MM/dd');
      }
      
      weeks.push({
        weekKey: weekKey,
        weekLabel: weekLabel,
        startDate: weekStart,
        endDate: weekEnd,
        isCurrentWeek: i === 0
      });
    }
    
    // Process workout data
    data.forEach(function(row) {
      var workoutDate = new Date(row[0]); // Column A is Date
      var muscle = row[2] || 'Other';     // Column C is Muscle
      var repsString = (row[4] || '').toString(); // Column E is Reps
      
      // Count number of sets (number of comma-separated reps)
      var setsCount = repsString ? repsString.split(',').length : 0;
      
      // Find which week this workout belongs to
      weeks.forEach(function(week) {
        if (workoutDate >= week.startDate && workoutDate <= week.endDate) {
          if (!muscleGroups[muscle]) {
            muscleGroups[muscle] = {};
            weeks.forEach(function(w) {
              muscleGroups[muscle][w.weekKey] = 0;
            });
          }
          muscleGroups[muscle][week.weekKey] += setsCount;
        }
      });
    });
    
    // Sort muscle groups alphabetically
    var sortedMuscleGroups = {};
    Object.keys(muscleGroups).sort().forEach(function(key) {
      sortedMuscleGroups[key] = muscleGroups[key];
    });
    
    var result = {
      weeks: weeks, // Already in descending order (most recent first)
      muscleGroups: sortedMuscleGroups
    };
    
    Logger.log('Stats calculated for ' + Object.keys(muscleGroups).length + ' muscle groups');
    return JSON.stringify(result);
  } catch (error) {
    Logger.log('Error in getWorkoutStats: ' + error.toString());
    throw new Error('Failed to load workout stats: ' + error.toString());
  }
}
// ===== WORKOUT SUBMISSION =====
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
        weightGroups[weight].reps.push(Number(set.reps) || 0);
        // Convert RIR to number, handle empty values as empty string
        var rirValue = set.rir && set.rir !== '' ? Number(set.rir) : '';
        weightGroups[weight].rirs.push(rirValue);
      });
      
      // Create a separate row for each weight used
      Object.keys(weightGroups).forEach(function(weight) {
        var groupData = weightGroups[weight];
        var repsString = groupData.reps.join(',');
        // Filter out empty RIR values and join
        var validRirs = groupData.rirs.filter(function(rir) { return rir !== ''; });
        var rirsString = validRirs.length > 0 ? validRirs.join(',') : '';
        
        // Prepare row data
        // Format: Date | Exercise | Muscle | Weight | Reps | RIR
        rowsToInsert.push([
          dateOnly,        // Date (MM/dd/yyyy format)
          exercise,        // Exercise name only
          muscle,          // Muscle group
          Number(weight),  // Weight (single value for this row)
          repsString,      // Reps for this weight (comma-separated)
          rirsString       // RIR ratings for this weight (comma-separated, numbers only)
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

