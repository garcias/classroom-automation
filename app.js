// TODO
// ========
// [ ] Minimize scopes requested
// [x] Create many assignments in a batch
// [ ] Generate a report about student completion/grades on assignments
// [x] Named sheets are refreshed not deleted + created, with option to delete
// [x] Allow to set name of sheet in sheet-factory functions
// [ ] Generalize some functions to library
// [ ] Record some tricks for object manipulation into Quiver
// [x] Source control
// [x] copy_content embeds url to the file in the title
// [x] Refresh submissions list
// [x] Automatically refresh submissions list after merge
// [ ] Test do_grade_completion to see if patching works

function myFunction() {
  // The comment below will trigger authorization dialog (yes, even as a comment)
  // ref: https://stackoverflow.com/questions/42796630/
  // Classroom.Courses.Coursework.remote( course_id )
  // Classroom.Courses.Coursework.create( course_id )

  course_id = '535370341314';  // KEEP test course
  batch_sheet = SpreadsheetApp.getActive().getSheetByName('batch');
}

function examples() {
  // clear_log();
  // var sheet = SpreadsheetApp.getActiveSheet();
  // var sheet = SpreadsheetApp.getActive().insertSheet('test');
  // sheet.getRange( 1, 1, 1, 3 ).setValues( [[ 'course', 'assignment', 'submission' ]] );
  // log_objects( list_courses(), "list_courses()" );
  // log_objects( list_students( '453549119831') );
  // log_objects( list_assignments_full( '453549119831' ) );
  // var submissions = list_submissions( '453549119831', '453612187654' );  // Week 12 preparation in Chem 126
  // log_objects( submissions, "list_submissions( '453549119831', '453612187654' )" );  
  // Logger.log( JSON.stringify(submissions, null, 2) );
  // courses = list_courses().map( course => `${course.id}: ${course.name}` ).slice(0,8);
}

function cleanup_journals() {
  course_id = '535370341314';
  batch_sheet = SpreadsheetApp.getActive().getSheetByName('batch');
  rows = read_sheet_to_objects( batch_sheet );
  topics = rows.filter( row => row.include ).map( row => spec_journal(row).topicId );
  unique = [ ... new Set( topics ) ];
  Logger.log(unique);
  unique.forEach( topic_id => { 
    response = Classroom.Courses.CourseWork.list( course_id, { courseWorkStates: ['DRAFT', 'PUBLISHED'] } );
    assignment_ids = response.courseWork.filter( a => a.topicId == topic_id  ).map( a => a.id);
    assignment_ids.forEach( id => { Classroom.Courses.CourseWork.remove( course_id, id ); })
    Classroom.Courses.Topics.remove( course_id, topic_id ) 
  });
}

// SCRIPTS

function onOpen(e) {
  // Classroom.Courses.Coursework.create( course_id )
  SpreadsheetApp.getUi()
    .createMenu('Classroom')
    .addItem('Refresh course list', 'do_refresh_course_list' )
    .addItem('Refresh assignments list', 'do_refresh_assignments_list' )
    .addItem('Refresh submissions list', 'do_refresh_submissions_list' )
    .addItem('Merge submissions', 'do_merge_submissions' )
    .addItem('Grade timely completion', 'do_grade_completion' )
      // This function only works on assignments created by this spreadsheet script.
    .addItem('Batch assign journals', 'do_setup_journals' )
    .addItem('Add batch rows', 'do_add_batch_rows' )
    .addToUi();
  console.info("UI built")
}

function do_grade_completion() {
  var active = SpreadsheetApp.getActive();
  var sheet = active.getSheetByName( 'submissions' ) || active.insertSheet( 'submissions' );

  current_data = read_sheet_to_objects( sheet );
  include_array = current_data.map( row => { return [ row.id, row.include ] } );
  include_map = Object.fromEntries( include_array );
  
  try {
    var assignment_sheet = SpreadsheetApp.getActive().getSheetByName('assignments');
    var assignments = read_sheet_to_objects( assignment_sheet );
    var selected_assignments = assignments.filter( assignment => assignment.include );

    var submissions = selected_assignments.map( a => list_submissions( a.courseId, a.id ) ).flat(1);
    if ( submissions.length > 0 ) {
      var timely_submissions = submissions.filter( 
        submission => !(submission.late) && submission.state == 'TURNED_IN'
      );
      responses = timely_submissions.map( submission => {
        return Classroom.Courses.CourseWork.StudentSubmissions.patch( { draftGrade: submission.maxPoints }, 
          submission.courseId, submission.courseWorkId, submission.id, { updateMask: 'draftGrade' }
        );
      });
      Logger.log( responses );
    } else {
      SpreadsheetApp.getUi().alert( `Selected assignments do not have submissions.` );
    }
  } catch(e) {
    SpreadsheetApp.getUi().alert( `${e}. Could not access submissions for selected assignments.` );
  }
}

function do_refresh_course_list() {  // modifies or creates sheet named 'courses'
  var active = SpreadsheetApp.getActive();
  var sheet = active.getSheetByName( 'courses' ) || active.insertSheet( 'courses' );

  // Try to remember which were selected for inclusion before
  // If the sheet doesn't exist, all this will be null and include will be false for every row
  current_data = read_sheet_to_objects( sheet );
  include_array = current_data.map( row => { return [ row.id, row.include ] } );
  include_map = Object.fromEntries( include_array );
  
  try {
    var course_list = list_courses().map( 
      course => Object.assign( course, { include: include_map[course.id] || false } ) 
    );  // add the property "include" for selection
  } catch(e) {
    SpreadsheetApp.getUi().alert( `${e}. Could not access course list.` );
  }

  var sheet = output_objects( course_list, sheet );

  // set checkbox validation on the last column "include"
  last_column = sheet.getLastColumn();
  number_rows = course_list.length;
  var rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sheet.getRange( 2, last_column, number_rows, 1 ).setDataValidation( rule );
  active.setActiveSheet( sheet );
}

function do_refresh_assignments_list() {
  var active = SpreadsheetApp.getActive();
  var sheet = active.getSheetByName( 'assignments' ) || active.insertSheet( 'assignments' );

  current_data = read_sheet_to_objects( sheet );
  include_array = current_data.map( row => { return [ row.id, row.include ] } );
  include_map = Object.fromEntries( include_array );
  
  try {
    var course_sheet = SpreadsheetApp.getActive().getSheetByName('courses');
    var courses = read_sheet_to_objects( course_sheet );
    var selected_courses = courses.filter( course => course.include );
    // selected_courses_ids = selected_courses.map( course => course.id );
    assignments = selected_courses.map( course => list_assignments_all( course.id ) ).flat(1);
    if ( assignments.length > 0 ) {
      var assignments_list = assignments.map( 
        assignment => Object.assign( assignment, { include: include_map[assignment.id] || false } )
      );  // add the property "include" for selection
      sheet = output_objects( assignments_list, sheet );

      //set checkbox validation on the last column "include"
      last_column = sheet.getLastColumn();
      number_rows = assignments.length;
      var rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
      sheet.getRange( 2, last_column, number_rows, 1 ).setDataValidation( rule );
      active.setActiveSheet( sheet );
    } else {
      SpreadsheetApp.getUi().alert( `Selected courses do not contain assignments.` );
    }
  } catch(e) {
    SpreadsheetApp.getUi().alert( `${e}. Could not access assignments for selected courses.` );
  }
}

function do_refresh_submissions_list() {
  var active = SpreadsheetApp.getActive();
  var sheet = active.getSheetByName( 'submissions' ) || active.insertSheet( 'submissions' );

  current_data = read_sheet_to_objects( sheet );
  include_array = current_data.map( row => { return [ row.id, row.include ] } );
  include_map = Object.fromEntries( include_array );
  
  try {
    var assignment_sheet = SpreadsheetApp.getActive().getSheetByName('assignments');
    var assignments = read_sheet_to_objects( assignment_sheet );
    var selected_assignments = assignments.filter( assignment => assignment.include );

    submissions = selected_assignments.map( a => list_submissions( a.courseId, a.id ) ).flat(1);
    if ( submissions.length > 0 ) {
      var submissions_list = submissions.map( 
        submission => Object.assign( submission, { include: include_map[submission.id] || false } )
      );  // add the property "include" for selection
      sheet = output_objects( submissions_list, sheet );

      //set checkbox validation on the last column "include"
      last_column = sheet.getLastColumn();
      number_rows = submissions.length;
      var rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
      sheet.getRange( 2, last_column, number_rows, 1 ).setDataValidation( rule );
      active.setActiveSheet( sheet );
    } else {
      SpreadsheetApp.getUi().alert( `Selected assignments do not have submissions.` );
    }
  } catch(e) {
    SpreadsheetApp.getUi().alert( `${e}. Could not access submissions for selected assignments.` );
  }
}

function do_merge_submissions() {
  var active = SpreadsheetApp.getActive();
  var sheet = active.getSheetByName( 'merge' ) || active.insertSheet( 'merge' );

  try {
    var assignment_sheet = SpreadsheetApp.getActive().getSheetByName('assignments');
    assignments = read_sheet_to_objects( assignment_sheet );
    selected_assignments = assignments.filter( assignment => assignment.include );
    target_urls = selected_assignments.map( assignment => ( { 
      id: assignment.id, title: assignment.title, url: assignment.alternateLink,
      mergedoc_url: merge_submissions( assignment.courseId, assignment.id )
    }));
    sheet = output_objects( target_urls, sheet );
    do_refresh_submissions_list();
    active.setActiveSheet( sheet );
  } catch(e) {
    SpreadsheetApp.getUi().alert( `Could not find submissions for selected assignment because of error: ${e}.`);
  }
}

function do_setup_journals() {
  // batch creates assignments, specified in sheet "batch", and marked as "include"
  // expects rows to specify courseId, topic name, title, schedule date/time
  // optionally expects due date/time, points, fileId of material to attach, description
  batch_sheet = SpreadsheetApp.getActive().getSheetByName('batch');
  rows = read_sheet_to_objects( batch_sheet ).filter( row => row.include );
  rows.forEach( row => { get_topic_id( row.courseId, row.topic ) });
  responses = rows.map( row => create_assignment(row) );
  batch_sheet.getRange( 'C2:C' ).setValue( false );
}

function do_add_batch_rows() {
  var ui = SpreadsheetApp.getUi();
  var num = ui.prompt( 'How many rows to add?').getResponseText();
  num = isNaN(num) ? 1 : Math.ceil(num);
  var batch_sheet = SpreadsheetApp.getActive().getSheetByName('batch');
  batch_sheet.insertRowsBefore( 2, num );
}


// ASSIGNMENT CREATION AND SPECIFICATION FUNCTIONS

function create_material_drive( course_id, title, topic_name, description, material_id ) {
  material_spec = {
    title: title, topicId: get_topic_id( course_id, topic_name ), description: description, state: 'DRAFT',
    materials: [{ driveFile: { shareMode: "VIEW", driveFile: { id: material_id, title: title }} }]
  };
  response = Classroom.Courses.CourseWorkMaterials.create( material_spec, course_id)
  return response;
}

function get_topic_id( course_id, topic_name ) {  // creates the topic if it doesn't exist
  let topic_list = Classroom.Courses.Topics.list( course_id ).topic;
  let existing_topic = topic_list ? topic_list.filter( t => t.name == topic_name ) : [];
  if ( existing_topic.length > 0 ) {
    topic_id = existing_topic[0].topicId;
  } else {
    response = Classroom.Courses.Topics.create( {name: topic_name }, course_id );
    topic_id = response.topicId;
  }
  return topic_id;
}

// function get_topic_id( course_id, topic_name ) {  // creates the topic if it doesn't exist
//   existing_topic = Classroom.Courses.Topics.list( course_id ).topic.filter( t => t.name == topic_name );
//   if ( existing_topic.length > 0 ) {
//     topic_id = existing_topic[0].topicId;
//   } else {
//     response = Classroom.Courses.Topics.create( {name: topic_name }, course_id );
//     topic_id = response.topicId;
//   }
//   return topic_id;
// }

function create_assignment( row ) {
  spec = spec_journal( row );
  let response = Classroom.Courses.CourseWork.create( spec, spec.courseId );
  Logger.log( spec.title );
  return response;
}

function spec_journal( row ) {
// convert from sheet-specified format to assignment-spec object that conforms to API spec
// new version requires columns for date and time, must be formatted as Sheets datetime <==> JS Date
  
  course_id = row.courseId.toString();
  topic_name = row.topic;
  points = row.points;
  materials_id = row.material.toString();
  
  existing_topic = Classroom.Courses.Topics.list( course_id ).topic.filter( t => t.name == topic_name );
  if ( existing_topic.length > 0 ) {
    topic_id = existing_topic[0].topicId;
  } else {
    throw `Topic "${topic_name}" does not exist for courseId ${course_id}.`
  }
  
  var due_datetime = row.due_date;
  due_datetime.setHours(   row.due_time.getHours() );
  due_datetime.setMinutes( row.due_time.getMinutes() );
  due_datetime.setSeconds( row.due_time.getSeconds() );

  let due_date = format_due_date( due_datetime );
  let due_time = format_due_time( due_datetime );

  var sch_datetime = row.sch_date;
  sch_datetime.setHours(   row.sch_time.getHours() );
  sch_datetime.setMinutes( row.sch_time.getMinutes() );
  sch_datetime.setSeconds( row.sch_time.getSeconds() );
  sch_datetime = format_sch_datetime( sch_datetime );

  points = points ? points : undefined;  // in case points is blank
  materials_spec = materials_id ? 
    [ { driveFile: { driveFile: { id: materials_id }, shareMode: 'STUDENT_COPY'} } ] : 
    undefined;  // in case material is blank

  return {
    state: 'DRAFT', workType: 'ASSIGNMENT', title: row.title, maxPoints: points, topicId: topic_id,
    dueDate: due_date, dueTime: due_time, scheduledTime: sch_datetime,
    materials: materials_spec, description: row.description, courseId: course_id,
  };
}

function format_sch_datetime( sch_datetime ) {  // formatted as ISO string
  return sch_datetime.toISOString();  
}

function format_due_date( due_datetime ) {  // returns { year: , month: , day: }
  if ( due_datetime ) {
    due_date = {year: due_datetime.getUTCFullYear(), month: due_datetime.getUTCMonth() + 1, day: due_datetime.getUTCDate()};
  } else {  // in case due date is blank
    due_date = undefined;
  }
  return due_date;
}

function format_due_time( due_datetime ) {  // returns { hours: , minutes: } in UTC
  if ( due_datetime ) {
    due_time = { hours: due_datetime.getUTCHours(), minutes: due_datetime.getUTCMinutes() };
  } else {  // in case due date is blank
    due_time = undefined;   
  }
  return due_time;
}

// DOCUMENT MANIPULATION FUNCTIONS

function merge_submissions( course_id, assignment_id, limit=undefined ) {  // returns url of new document
  // sort and create array of { email:, fileId: }
  var submissions = list_submissions( course_id, assignment_id );
  sorted = submissions.map( s => [s.email, s.fileId] ).sort();
  submissions = sorted.map( tuple => { return { email: tuple[0], fileId: tuple[1] } } ).slice( 0, limit );

  title = Classroom.Courses.CourseWork.get( course_id, assignment_id).title;
  var merge_doc = DocumentApp.create( `merge submissions for ${title}` );
  var merge_doc_id = merge_doc.getId();
  merge_doc.saveAndClose();
  
  for ( let submission of submissions ) {
    var merge_doc = DocumentApp.openById( merge_doc_id );
    var target = merge_doc.getBody();

    var file = DriveApp.getFileById( submission.fileId );
    if( file.getMimeType() == MimeType.GOOGLE_DOCS ) {
      var doc = DocumentApp.openById( submission.fileId );
      var source = doc.getBody();

      // Indicate email in the title; original doc name in subtitle
      owner = submission.email;
      target.appendParagraph( owner ).setHeading(DocumentApp.ParagraphHeading.TITLE);
      doc_title = doc.getName();
      target.appendParagraph( doc_title )
        .setHeading( DocumentApp.ParagraphHeading.SUBTITLE )
        .setLinkUrl( doc.getUrl() );
      Logger.log( `${owner}: ${doc_title}`)
      replaceFootnotes( source );
      copyContent( source, target );      
    } else {
      // Indicate email in the title
      owner = submission.email;
      target.appendParagraph( owner ).setHeading(DocumentApp.ParagraphHeading.TITLE);
      target.appendParagraph( file.getName() )
        .setHeading( DocumentApp.ParagraphHeading.SUBTITLE )
        .setLinkUrl( file.getUrl() );
      warning = '\n>>>>> Source file is not a Google Doc; could not copy content. <<<<<\n'
      target.appendParagraph( warning );
    }
    merge_doc.saveAndClose();
  }
  return merge_doc.getUrl();
}

function copyContent( source_body, target_body ) {
  var num_elements = source_body.getNumChildren();
  for (var i=0; i < num_elements; i++ ) {
    e = source_body.getChild(i).copy();
    e_type = e.getType();
    if( e.getAttributes().HEADING == DocumentApp.ParagraphHeading.TITLE ) {
      continue;
    }
    try{
      switch ( e_type ) {
        case DocumentApp.ElementType.HORIZONTAL_RULE:
          target_body.appendHorizontalRule();
          break;
        case DocumentApp.ElementType.INLINE_IMAGE:
          target_body.appendImage(e);
          break;
        case DocumentApp.ElementType.LIST_ITEM:
          target_body.appendListItem(e);
          break;
        case DocumentApp.ElementType.PAGE_BREAK:
          target_body.appendPageBreak();
          break;
        case DocumentApp.ElementType.PAGE_BREAK:
          target_body.appendPageBreak(e);
          break;
        case DocumentApp.ElementType.PARAGRAPH:
          target_body.appendParagraph(e);
          break;
        case DocumentApp.ElementType.TABLE:
          target_body.appendTable(e);
          break;
        default:
          console.warn( 'Unsupported element: not copied');
          target_body.appendParagraph(`\n>>>>> Warning: Could not copy unknown element from source. <<<<<\n`);
      }  
    } catch(e) {
      console.warn( `Failed to copy: ${e_type}` );
      console.warn(e);
      target_body.appendParagraph(`\n>>>>> Warning: Could not copy ${e_type} from source. <<<<<\n`);
      target_body.appendParagraph(`\n>>>>> Warning: Refer to original source. <<<<<\n`);
      // console.warn( source_body.getChild(i).getNumChildren() );
    }
  } 
}

function replaceFootnotes( source_body ) {
  var searchType = DocumentApp.ElementType.FOOTNOTE;
  var searchResult = null;
  
  while (searchResult = source_body.findElement(searchType, searchResult)) {
    var footnote = searchResult.getElement().asFootnote();
    var footnote_parent = footnote.getParent();
    var footnote_index = footnote_parent.getChildIndex(footnote);
    var footnote_text = footnote.getFootnoteContents().getText();
    footnote_parent.insertText( footnote_index + 1, " [[ "+footnote_text+" ]] " );
    // footnote.removeFromParent();
  }
  
  var searchResult = null;
  while (searchResult = source_body.findElement(searchType)) {
    var footnote = searchResult.getElement().asFootnote();
    footnote.removeFromParent();
  }
}

// BASIC LISTING FUNCTIONS, return object or [object]

function list_students( course_id ) { // returns hash table of {id: {email: , name: }}
  var response = Classroom.Courses.Students.list( course_id );
  students = response.students.map( student => {
    return { 
      id: student.profile.id, profile: { email: student.profile.emailAddress, name: student.profile.name.fullName } 
    };
  }).reduce( 
    (target,student) => { target[student.id] = student.profile; return target }, {} 
  );
  return students;
}

function list_courses() { // returns array of { id, section:, courseState:, alternateLink:, name: }
  var response = Classroom.Courses.list();
  var courses = response.courses;
  desired_keys = [ 'id', 'section', 'courseState', 'alternateLink', 'name' ]
  return courses.map( course => filter_keys( course, desired_keys ) );
}

function list_assignments_all( course_id ) { 
  // returns array of { id:, title:, state:, topicId:, description:, materials:, maxPoints: }
  var response = Classroom.Courses.CourseWork.list( course_id, { courseWorkStates: ['DRAFT', 'PUBLISHED'] } );
  var assignments = response.courseWork.map( a => {
    let d = a.dueDate ?
      new Date( Date.UTC(a.dueDate.year, a.dueDate.month - 1, a.dueDate.day, a.dueTime.hours, 0) ) : undefined;
    let date = d ? `${d.toString().slice(0,3)} ${d.toISOString().slice(0,10)} ${d.toTimeString().slice(0,5)}` : undefined;
    return Object.assign( a, { due: date } );
  });
  desired_keys = [ 
    'id', 'title', 'alternateLink', 'state', 'courseId', 'topicId', 'description', 
    'materials', 'maxPoints', 'due' 
  ];
  let curated = assignments.map( assignment => filter_keys( assignment, desired_keys ) );
  return curated;
}

function list_assignments_full( course_id ) {
  var response = Classroom.Courses.CourseWork.list( course_id );
  return response.courseWork;
}

function list_submissions( course_id, assignment_id ) {  // returns array of objects, one per *attachment*
  var students = list_students( course_id );
  var assignment = Classroom.Courses.CourseWork.get( course_id, assignment_id);
  var title = assignment.title;
  var max = assignment.maxPoints;
  var response = Classroom.Courses.CourseWork.StudentSubmissions.list( course_id, assignment_id);
  submissions = response.studentSubmissions.map(
    submission => submission.assignmentSubmission.attachments.map(
      attachment => { return {
          id: submission.id,
          courseId: submission.courseId,
          courseWorkId: submission.courseWorkId,
          assignment: title,
          userId: submission.userId,
          email: students[submission.userId].email,
          fileId: attachment.driveFile.id,
          title: attachment.driveFile.title,
          url: attachment.driveFile.alternateLink,
          state: submission.state, late: submission.late,
          maxPoints: max,
          draftGrade: submission.draftGrade, assignedGrade: submission.assignedGrade,
        }
      }
    ) 
  )
  return submissions.flat();
}

function reset_coursework( course_id ) {
  var response = Classroom.Courses.CourseWork.list( course_id );
  var assignments = response.courseWork;
  assignments.forEach( assignment => {
    r = Classroom.Courses.CourseWork.remove( course_id, assignment.id );
    Logger.log(r);
  })
}

function filter_keys( source_object, keys_array ) {  // returns object with subset of properties
  return keys_array.reduce( (obj, key) => { obj[key] = source_object[key]; return obj}, {} )
}


// SHEET MANIPULATION FUNCTIONS

function read_sheet_to_objects( sheet ) {  // sheet organized in table format; returns array of objects
  var rows = sheet.getDataRange().getValues();
  var keys = rows.splice(0,1)[0];  // side effect: rows loses the first item
  var data = rows.map( row => 
    Object.assign( {}, ...keys.map( (key,i) => ( { [key]: row[i] }) ) ) 
  );
  return data;
}

function output_arrays( data, sheet = undefined ) {  // data is an array of objects; returns an unnamed Sheet
  if ( data.length < 1 ) throw 'output_arrays: array is empty';
  if ( sheet == null ) {
    sheet = SpreadsheetApp.getActive().insertSheet();
  }
  num_rows = data.length;
  num_cols = data[0].length;
  // sheet.insertRowsAfter(1, num_rows);
  sheet.getDataRange().clearContent();
  insert_range = sheet.getRange( 1, 1, num_rows, num_cols );
  insert_range.setValues( data );
  clear_empty_rows_and_columns( sheet );
  return sheet;
}

function output_objects( data, sheet = undefined ) {  // data is an array of objects; returns an unnamed Sheet
  if ( data.length < 1 ) throw 'output_objects: array of objects is empty';
  headers = Object.keys( data[0] );
  var body = data.map( row => Object.values( row ) )
  new_arr = [headers].concat(body);
  var sheet = output_arrays( new_arr, sheet );
  sheet.setFrozenRows(1);
  return sheet;
}

function log_arrays( data, mesg ) {  // data is an array of arrays
  var active = SpreadsheetApp.getActive();
  var sheet = active.getSheetByName( 'log-output' ) || active.insertSheet( 'log-output' );
  num_rows = data.length;
  num_cols = data[0].length;
  sheet.insertRowsAfter(1, num_rows + 1);
  insert_range = sheet.getRange( 2, 1, num_rows, num_cols );
  insert_range.setValues( data );
  sheet.insertRowsAfter(1, 1);
  sheet.getRange(2,1,1,1).getCell( 1, 1 ).setValue( `CMD: ${mesg}` );
  clear_empty_rows_and_columns( sheet );
}

function log_objects( data, mesg ) {  // data is an array of objects
  var active = SpreadsheetApp.getActive();
  var sheet = active.getSheetByName( 'log-output' ) || active.insertSheet( 'log-output' );
  headers = Object.keys( data[0] );
  num_cols = headers.length;
  sheet.insertRowsAfter( 1, 1 );
  insert_range = sheet.getRange( 2, 1, 1, num_cols );
  insert_range.setValues( [ headers ] );

  num_rows = data.length;
  arr = data.map( obj => headers.map( header => obj[header] ) );
  console.log(arr);
  sheet.insertRowsAfter( 2, num_rows + 1);
  insert_range = sheet.getRange( 3, 1, num_rows, num_cols );
  insert_range.setValues( arr );
  sheet.insertRowsAfter(1, 1);
  sheet.getRange(2,1,1,1).getCell( 1, 1 ).setValue( `CMD: ${mesg}` );
  clear_empty_rows_and_columns( sheet );
}

function clear_log() {
  var active = SpreadsheetApp.getActive();
  var sheet = active.getSheetByName( 'log-output' ) || active.insertSheet( 'log-output' );
  last_row_pos = sheet.getFrozenRows() + 1;
  target_rows = sheet.getMaxRows() - last_row_pos;
  if ( target_rows > 0 ) { sheet.deleteRows( last_row_pos + 1, target_rows ) }
  sheet.getRange(1,1,1,1).getCell( 1, 1 ).setValue( "OUTPUT LOG" );
  sheet.getRange(2,1,1,1).getCell( 1, 1 ).setValue( "CMD: clear_log()" );
  sheet.setFrozenRows(1);
}

function clear_empty_rows_and_columns( sheet ) {
  last_row_pos = sheet.getLastRow();
  target_rows = sheet.getMaxRows() - last_row_pos;
  if ( target_rows > 0 ) { sheet.deleteRows( last_row_pos + 1, target_rows ) }
  last_col_pos = sheet.getLastColumn();
  target_cols = sheet.getMaxColumns() - last_col_pos;
  if ( target_cols > 0 ) { sheet.deleteColumns( last_col_pos + 1, target_cols ) }
}

function clear_all_rows() {
  var sheet = SpreadsheetApp.getActiveSheet();
  last_row_pos = sheet.getFrozenRows() + 1;
  max_rows = sheet.getMaxRows();
  sheet.deleteRows( last_row_pos + 1, max_rows - last_row_pos );
}
