// TODO
// ========
// [ ] Allow to set location of new merge file
// [ ] Minimize scopes requested
// [x] Create many assignments in a batch
// [ ] Generate a report about student completion/grades on assignments
// [ ] Named sheets are refreshed not deleted + created, with option to delete
// [x] Allow to set name of sheet in sheet-factory functions
// [ ] Generalize some functions to library
// [ ] Record some tricks for object manipulation into Quiver
// [x] Source control
// [x] copy_content embeds url to the file in the title

var TZ = 4; // local time zone offset from GMT

function myFunction() {
  // The comment below will trigger authorization dialog (yes, even as a comment)
  // ref: https://stackoverflow.com/questions/42796630/
  // Classroom.Courses.Coursework.remote( course_id )
  // Classroom.Courses.Coursework.create( course_id )

  course_id = '535370341314';  // KEEP test course
  batch_sheet = SpreadsheetApp.getActive().getSheetByName('batch');
  SpreadsheetApp.getActive().setActiveSheet( batch_sheet );
}

function examples() {
  // var sheet = SpreadsheetApp.getActiveSheet();
  // var sheet = SpreadsheetApp.getActive().getSheetByName("courselist").activate();
  // clear_log();
  // log_objects( list_courses(), "list_courses()" );
  // log_objects( list_assignments( '496710505024' ), "list_assignments" ); // Chem 126 M pm section
  // log_objects( list_assignments_all( '496710505024' ) ); // Chem 126 M pm section
  // to_delete = ["496710505073"];
  // to_delete.forEach( id => {
  //   Classroom.Courses.CourseWork.remove( '496710505024', id );
  // })
  // Logger.log( to_delete );
  // var submissions = list_submissions( '453549119831', '453612187654' );  // Week 12 preparation in Chem 126
  // log_objects( submissions, "list_submissions( '453549119831', '453612187654' )" );
  // new_doc_url = merge_submissions( '453549119831', '453612187654' );
  // log_arrays( [[ "merge doc url", new_doc_url ]], "merge_submissions( '453549119831', '453612187654'" );
  
  // Logger.log( JSON.stringify(submissions, null, 2) );
  // log_objects( list_students( '453549119831') );
  // log_objects( list_assignments_full( '453549119831' ) );
  // var sheet = SpreadsheetApp.getActive().getSheetByName('merge');
  // sheet.getRange( 1, 1, 1, 3 ).setValues( [[ 'course', 'assignment', 'submission' ]] );
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
    .addItem('Merge submissions ...', 'do_merge_submissions' )
    .addToUi();
  console.info("UI built")
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
    selected_courses_ids = selected_courses.map( course => course.id );
    assignments = selected_courses_ids.map( id => list_assignments_all( id ) ).flat(1);
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

function do_merge_submissions() {
  var active = SpreadsheetApp.getActive();
  var sheet = active.getSheetByName( 'merge' ) || active.insertSheet( 'merge' );

  // current_data = read_sheet_to_objects( sheet );
  // include_array = current_data.map( row => { return [ row.id, row.include ] } );
  // include_map = Object.fromEntries( include_array );

  try {
    var assignment_sheet = SpreadsheetApp.getActive().getSheetByName('assignments');
    assignments = read_sheet_to_objects( assignment_sheet );
    selected_assignments = assignments.filter( assignment => assignment.include );
    target_urls = selected_assignments.map( assignment => ( { 
      id: assignment.id, title: assignment.title,
      url: merge_submissions( assignment.courseId, assignment.id )
    }));
    sheet = output_objects( target_urls, sheet );
    active.setActiveSheet( sheet );
  } catch(e) {
    SpreadsheetApp.getUi().alert( `Could not find submissions for selected assignment because of error: ${e}.`);
  }
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
  existing_topic = Classroom.Courses.Topics.list( course_id ).topic.filter( t => t.name == topic_name );
  if ( existing_topic.length > 0 ) {
    topic_id = existing_topic[0].topicId;
  } else {
    response = Classroom.Courses.Topics.create( {name: topic_name }, course_id );
    topic_id = response.topicId;
  }
  return topic_id;
}

function setup_journals( ) {
  // batch creates assignments, specified in sheet "batch", and marked as "include"
  // expects rows to specify courseId, topic name, title, schedule date/time
  // optionally expects due date/time, points, fileId of material to attach, description
  batch_sheet = SpreadsheetApp.getActive().getSheetByName('batch');
  rows = read_sheet_to_objects( batch_sheet ).filter( row => row.include );
  rows.forEach( row => {
    existing_topics = Classroom.Courses.Topics.list( row.courseId ).topic.map( t => t.name );
    if( !existing_topics.includes( row.topic ) ) {
      Classroom.Courses.Topics.create( {name: row.topic }, row.courseId );
    }
  })
  
  responses = rows.map( row => create_assignment(row) );
}

function create_assignment( row ) {
  spec = spec_journal( row );
  let response = Classroom.Courses.CourseWork.create( spec, spec.courseId );
  Logger.log( spec.title );
  return response;
}

function spec_journal( row ) {
// convert from sheet-specified format to assignment-spec object that conforms to API spec
  
  course_id = row.courseId.toString();
  topic_name = row.topic;
  points = row.points;
  materials_id = row.material.toString();
  // in local time
  sch_year = row.sch_year; sch_month = row.sch_month; sch_day = row.sch_day, sch_hour = row.sch_hour; sch_min = 0;
  due_year = row.due_year; due_month = row.due_month; due_day = row.due_day, due_hour = row.due_hour; due_min = 0;
  
  existing_topic = Classroom.Courses.Topics.list( course_id ).topic.filter( t => t.name == topic_name );
  if ( existing_topic.length > 0 ) {
    topic_id = existing_topic[0].topicId;
  } else {
    throw `Topic ${topic_name} does not exist for courseId ${course_id}.`
  }

  sch_datetime = format_sch_datetime( sch_year, sch_month, sch_day, sch_hour, sch_min );
  due_date = format_due_date( due_year, due_month, due_day, due_hour, due_min );
  due_time = format_due_time( due_year, due_month, due_day, due_hour, due_min );

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

function format_sch_datetime( sch_year, sch_month, sch_day, sch_hour, sch_min ) {  // formatted as ISO string
  sch_datetime = new Date( sch_year, sch_month - 1, sch_day, sch_hour, sch_min );
  return sch_datetime.toISOString();  
}

function format_due_date( due_year, due_month, due_day, due_hour, due_min ) {  // returns { year: , month: , day: }
  due_datetime = new Date( due_year, due_month - 1, due_day, due_hour + TZ, due_min );
  if ( due_year && due_month && due_day ) {
    due_date = {year: due_datetime.getFullYear(), month: due_datetime.getMonth() + 1, day: due_datetime.getDate()};
  } else {  // in case due date is blank
    due_date = undefined;
  }
  return due_date;
}

function format_due_time( due_year, due_month, due_day, due_hour, due_min ) {  // returns { hours: , minutes: }
  due_datetime = new Date( due_year, due_month - 1, due_day, due_hour + TZ, due_min );
  if ( due_year && due_month && due_day ) {
    due_time = { hours: due_datetime.getHours(), minutes: due_datetime.getMinutes() };
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
  var target = merge_doc.getBody();

  for ( let submission of submissions ) {
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

function list_assignments( course_id ) { // returns array of { id:, title:, state:, topicId:, description:, materials: }
  var response = Classroom.Courses.CourseWork.list( course_id );
  var assignments = response.courseWork;
  desired_keys = [ 'id', 'title', 'state', 'courseId', 'topicId', 'description', 'materials' ]
  return assignments.map( assignment => filter_keys( assignment, desired_keys ) );
}

function list_assignments_all( course_id ) { // returns array of { id:, title:, state:, topicId:, description:, materials: }
  var response = Classroom.Courses.CourseWork.list( course_id, { courseWorkStates: ['DRAFT', 'PUBLISHED'] } );
  var assignments = response.courseWork;
  desired_keys = [ 'id', 'title', 'state', 'courseId', 'topicId', 'description', 'materials' ]
  return assignments.map( assignment => filter_keys( assignment, desired_keys ) );
}

function list_assignments_full( course_id ) {
  var response = Classroom.Courses.CourseWork.list( course_id );
  return response.courseWork;
}

function list_submissions( course_id, assignment_id ) {  // returns array of objects, one per *attachment*
  var students = list_students( course_id );
  var response = Classroom.Courses.CourseWork.StudentSubmissions.list( course_id, assignment_id);
  submissions = response.studentSubmissions.map(
    submission => submission.assignmentSubmission.attachments.map(
      attachment => { return {
          id: submission.id,
          userId: submission.userId,
          email: students[submission.userId].email,
          fileId: attachment.driveFile.id,
          title: attachment.driveFile.title,
          url: attachment.driveFile.alternateLink,
          courseWorkId: submission.courseWorkId,
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
