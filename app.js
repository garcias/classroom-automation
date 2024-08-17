// TODO
// ========
// [x] Update batch assignment
// [ ] Fix how submissions displays nested data
// [ ] Add script to collect all comments made on Google Docs submitted for a given assignment
// [ ] Add script to detect number of missing or late submissions per student in a course
// [ ] Use JSDoc to generate documentation page for GitHub repo
// [ ] Map the Resource-type classes to the Classroom API more consistently

Settings = {
  pageSize: 30,
}

function myFunction() {
  // The comment below will trigger authorization dialog (yes, even as a comment)
  // ref: https://stackoverflow.com/questions/42796630/
  // Classroom.Courses.Coursework.remote( course_id )
  // Classroom.Courses.Coursework.create( course_id )
}

// SCRIPTS FOR SHEETS INTERFACE

function onOpen(e) {
  // Classroom.Courses.Coursework.create( course_id )
  SpreadsheetApp.getUi()
    .createMenu('Classroom')
    .addItem('Refresh course list', 'do_refresh_course_list' )
    .addItem('Refresh assignments list', 'do_refresh_assignments_list' )
    .addItem('Refresh submissions list', 'do_refresh_submissions_list' )
    .addItem('Merge submissions', 'do_merge_submissions' )
      // This function only works on assignments created by this spreadsheet script.
    .addItem('Batch assign journals', 'do_batch_assign' )
    .addToUi();
  console.info("UI built")
}

function do_refresh_course_list() {  // modifies or creates sheet named 'courses'
  const active = SpreadsheetApp.getActive();
  const sheet = active.getSheetByName( 'courses' ) || active.insertSheet( 'courses' );
  
  try {
    const course_list = Course.list();
    let data = SheetTable.from( course_list );
    data.updateAndPreserve( sheet, 'id' )
  } catch(e) {
    SpreadsheetApp.getUi().alert( `${e}. Could not access course list.` );
  }  
}

function do_refresh_assignments_list() {
  const active = SpreadsheetApp.getActive();
  const sheet = active.getSheetByName( 'assignments' ) || active.insertSheet( 'assignments' );
    
  try {
    const course_sheet = SpreadsheetApp.getActive().getSheetByName('courses');
    const courses = SheetTable.fromSheet( course_sheet );
    const selected_courses = courses.filter( course => course.select );
    
    const assignments_per_course = selected_courses.map( course => {
      return Course.getById( course.id ).getAssignments();
    });
    const assignments = assignments_per_course.flat(1);
    if ( assignments.length > 0 ) {
      let data = SheetTable.from( assignments );
      data.updateAndPreserve( sheet, 'id' );
      active.setActiveSheet( sheet );
    } else {
      SpreadsheetApp.getUi().alert( `Selected courses do not contain assignments.` );
    }
  } catch(e) {
    SpreadsheetApp.getUi().alert( `${e}. Could not access assignments for selected courses.` );
  }
}

function do_refresh_submissions_list() {
  const active = SpreadsheetApp.getActive();
  const sheet = active.getSheetByName( 'submissions' ) || active.insertSheet( 'submissions' );
  
  try {
    const assignment_sheet = SpreadsheetApp.getActive().getSheetByName('assignments');
    const assignments = SheetTable.fromSheet( assignment_sheet );
    const selected_assignments = assignments.filter( assignment => assignment.select );

    const submissions_per_assignment = selected_assignments.map( assignment => {
      return Assignment.get( assignment.courseId, assignment.id ).getSubmissions();
    });
    const submissions = submissions_per_assignment.flat(1);
    if ( submissions.length > 0 ) {

      // Submission is a highly-nested structure, so need to map to a row in the table
      const submissions_map = submissions.map( submission => {
        const assignment = Assignment.get( submission.courseId, submission.courseWorkId );
        const student = Student.get( submission.courseId, submission.userId );
        const attachments = submission.assignmentSubmission.attachments ||
          [ { driveFile: { title: null, alternateLink: null } } ];  // in case there are no attachments
        Logger.log( submission.assignmentSubmission );
        return attachments.map( attachment => {
          const denested_fields = {
            assignment: assignment.title,
            email: student.getEmail(),
            maxPoints: assignment.maxPoints,
            title: attachment.driveFile.title,
            url: attachment.driveFile.alternateLink,
          }
          return Object.assign( submission, denested_fields );
        });
      });
      let data = SheetTable.from( submissions_map.flat(1) );
      data.updateAndPreserve( sheet, 'id' );
      active.setActiveSheet( sheet );
    } else {
      SpreadsheetApp.getUi().alert( `Selected assignments do not have submissions.` );
    }
  } catch(e) {
    SpreadsheetApp.getUi().alert( `${e}. Could not access submissions for selected assignments.` );
  }
}

function do_merge_submissions() {
  const active = SpreadsheetApp.getActive();
  const sheet = active.getSheetByName( 'merge' ) || active.insertSheet( 'merge' );

  try {
    const assignment_sheet = SpreadsheetApp.getActive().getSheetByName('assignments');
    const assignments = SheetTable.fromSheet( assignment_sheet );
    const selected_assignments = assignments.filter( assignment => assignment.select );

    merge_docs = selected_assignments.map( a => {
      let merge_doc = new MergeDoc( `Merge of ${a.title}` );
      const course = Course.getById( a.courseId );
      const assignment = course.getAssignment( a.id );
      const students = course.getStudents();
      const index = Object.fromEntries( students.map( s => [ s.userId, s.getEmail() ] ) );
      let submissions = assignment.getSubmissions();
      submissions.sort( (a,b) => {
        return index[a.userId] > index[b.userId] ? 1 : index[a.userId] < index[b.userId] ? -1 :0
      });
      submissions.forEach( submission => {
        submission.driveFiles.forEach( drivefile => {
          const owner = course.getStudent( submission.userId ).getEmail();
          merge_doc.addTitle( owner );
          merge_doc.addSubtitle( drivefile.title, drivefile.alternateLink );
          let file_type = DriveApp.getFileById( drivefile.id ).getMimeType();
          if ( file_type == MimeType.GOOGLE_DOCS ) {
            try {
              let doc = DocumentApp.openById( drivefile.id );
              merge_doc.addDocument( doc );
              Logger.log( `${submission.id}: ${drivefile.title}`)
            } catch(e) {
              console.warn( `Drive file ${drivefile.title} could not be added because "${e}"`);
              merge_doc.addMessage( `Drive file ${drivefile.title} could not be added because "${e}"` );
            }
          } else {
            console.warn( `Drive file ${drivefile.title} skipped because it's not a Google Doc` );
            merge_doc.addMessage( `Drive file ${drivefile.title} skipped because it's not a Google Doc` );
          }
        })
      });
      Logger.log( `submission driveFiles merged into ${ merge_doc.doc.getUrl() }`)
      return {
        id: assignment.id, title: assignment.title, 
        assignmentUrl: assignment.alternateLink,
        mergeUrl: merge_doc.doc.getUrl(),
      };
    });

    let data = SheetTable.from( merge_docs );
    data.updateAndPreserve( sheet, 'id' );
    active.setActiveSheet( sheet );
  } catch(e) {
    SpreadsheetApp.getUi().alert( `Could not find submissions for selected assignment because of error: ${e}.`);
  }
}

function do_batch_assign() {
  // batch creates assignments, specified in sheet "batch", and marked as "include"
  // expects rows to specify courseId, topic name, title, schedule date/time
  // optionally expects due date/time, points, fileId of material to attach, description
  const batch_sheet = SpreadsheetApp.getActive().getSheetByName('batch');
  const specs = SheetTable.fromSheet( batch_sheet );
  const selected_specs = specs.filter( spec => spec.select );

  const assignments = selected_specs.map( spec => {
    const journal = Journal.fromObject( spec );
    const course = Course.getById( spec.courseId );
    const new_assignment = course.createAssignment ( journal );
    Logger.log( new_assignment.title );
    return new_assignment;
  });

  const updated_specs = specs.map( spec => {
    return Object.assign( spec, { select: false } )
  });
  SheetTable.from( updated_specs ).update( batch_sheet );
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

class Journal {
  constructor( spec ) { Object.assign( this, spec ) }
  static fromObject( spec ) {
    const course = Course.getById( spec.courseId );
    const topic = course.getTopicByName( spec.topic );

    let dueJS = new DateTime( spec.due_date );
    if (spec.due_date && spec.due_time) {
      dueJS.setHours(   spec.due_time.getHours() );
      dueJS.setMinutes( spec.due_time.getMinutes() );
      dueJS.setSeconds( spec.due_time.getSeconds() );
    } else {
      dueJS = undefined;
    }

    let schJS = new DateTime( spec.sch_date );
    if (spec.sch_date && spec.sch_time) {
      schJS.setHours(   spec.sch_time.getHours() );
      schJS.setMinutes( spec.sch_time.getMinutes() );
      schJS.setSeconds( spec.sch_time.getSeconds() );
    } else {
      schJS = undefined;
    }

    const materials_id = spec.material.toString();
    const materials_spec = materials_id ? 
      [ { driveFile: { driveFile: { id: materials_id }, shareMode: 'STUDENT_COPY'} } ] : 
      undefined;  // in case material was blank

    return new this({ 
      topicId: topic.topicId, 
      title: spec.title,
      maxPoints: spec.points ? spec.points : undefined, 
      state: 'DRAFT', 
      description: spec.description,
      materials: materials_spec,
      dueDate: dueJS.asDate(), 
      dueTime: dueJS.asTime(), 
      scheduledTime: schJS.toISOString(),
    })
  }
}

// DOCUMENT MANIPULATION CLASSES
/** 
 * Handler for a Google Doc that can accumulate the content of other documents into itself
 * @example
 * const mergedoc = new MergeDoc( "All examples" );
 * mergedoc.addTitle( "Example1" );
 * mergedoc.addDocument( example1 );
 * mergedoc.addTitle( "Example2" );
 * mergedoc.addDocument( example2 );
 * mergedoc.addMessage( "End of examples" );
 */
class MergeDoc {
  /**
   * Create a new Google Doc to accumulate content and associate with a new Merger 
   * @param {String} title - desired title of created document
   */
  constructor( title ) {
    this.doc = DocumentApp.create( title );
    this.save();
  }

  /** Save content in the accumulator doc, close it, and reopen it */
  save() {
    let id = this.doc.getId();
    this.doc.saveAndClose();
    this.doc = DocumentApp.openById( id );
  }

  /** 
   * Copy content from source_doc into the accumulator doc
   * @param {Object} source_doc - an existing Google Doc to copy from
   */
  addDocument( source_doc ) {
    let target_doc = this.doc;
    let source_body = source_doc.getBody();
    let target_body = target_doc.getBody();
    this.replaceFootnotes( source_body );
    this.copyContent( source_body, target_body );
    this.save();
  }

  /**
   * Append text (styled as TITLE) to accumulator doc, usually to preface copied content
   * @param {String} title - desired title
   * @param {String} [url] - optional url for title text to link to
   */
  addTitle( title, url = undefined ) {
    let target_doc = this.doc;
    let target_body = target_doc.getBody();
    let text = target_body.appendParagraph( title ).setHeading( DocumentApp.ParagraphHeading.TITLE );
    if (url) {
      text.setLinkUrl( url );
    }
    this.save();
  }

  /**
   * Append text (styled as SUBTITLE) to accumulator doc, usually to preface copied content
   * @param {String} subtitle - desired subtitle
   * @param {String} [url] - optional url for title text to link to
   */
  addSubtitle( subtitle, url = undefined ) {
    let target_doc = this.doc;
    let target_body = target_doc.getBody();
    let text = target_body.appendParagraph( subtitle ).setHeading( DocumentApp.ParagraphHeading.SUBTITLE );
    if (url) {
      text.setLinkUrl( url );
    }
    this.save();
  }

  /**
   * Append text (styled as NORMAL) to accumulator doc, usually to record a message or warning
   * @param {String} message - desired message
   */
  addMessage( message ) {
    let target_doc = this.doc;
    let target_body = target_doc.getBody();
    let text = target_body.appendParagraph( message ).setHeading( DocumentApp.ParagraphHeading.NORMAL );
    this.save();
  }

  /**
   * Copies content from source_body to target_body
   * @param {Object} source_body - body of existing Google Doc to copy from
   * @param {Object} target_body - body of existing Google Doc to copy to
   * 
   * Recognizes common elements: horizontal rule, inline image, list item, page break, paragraph, table
   * Can handle postioned images in list item or paragraph
   * Casts code snippets to paragraph
   * Skips titles
   * Inserts warning message if element is not supported
   */
  copyContent( source_body, target_body ) {
    var num_elements = source_body.getNumChildren();
    for (let i=0; i < num_elements; i++ ) {
      let e = source_body.getChild(i).copy();
      let e_type = e.getType();
      if( e.getAttributes().HEADING == DocumentApp.ParagraphHeading.TITLE ) {
        continue;
      }
      try {
        switch ( e_type ) {
          case DocumentApp.ElementType.HORIZONTAL_RULE:
            target_body.appendHorizontalRule();
            break;
          case DocumentApp.ElementType.INLINE_IMAGE:
            target_body.appendImage(e);
            break;
          case DocumentApp.ElementType.LIST_ITEM:
            e.getPositionedImages().forEach( image => {
              let blob = image.getBlob();
              let last_index = e.getNumChildren();
              e.removePositionedImage( image.getId() );
              e.insertInlineImage( last_index, blob );
            });  // detached elements containing PositionedImages cannot be inserted
            target_body.appendListItem(e);
            break;
          case DocumentApp.ElementType.PAGE_BREAK:
            target_body.appendPageBreak();
            break;
          case DocumentApp.ElementType.PAGE_BREAK:
            target_body.appendPageBreak(e);
            break;
          case DocumentApp.ElementType.PARAGRAPH:
            e.getPositionedImages().forEach( image => {
              let width = image.getWidth();
              let height = image.getHeight();
              let blob = image.getBlob();
              let last_index = e.getNumChildren();
              e.removePositionedImage( image.getId() );
              e.insertInlineImage( last_index, blob ).setWidth(width).setHeight(height);
            });  // detached elements containing PositionedImages cannot be inserted
            target_body.appendParagraph(e);
            break;
          case DocumentApp.ElementType.TABLE:
            target_body.appendTable(e);
            break;
          default:
            if ( e_type.toString() == 'CODE_SNIPPET' ) {
              console.warn( `Casted ${e_type} into Paragraph of normal text` );
              target_body.appendParagraph(`\n>>>>> WARNING: Casted ${e_type} into Paragraph of normal text. <<<<<\n`);
              target_body.appendParagraph( e.asText().getText() );
              target_body.appendParagraph(`\n>>>>> WARNING: Refer to original source. <<<<<\n`);
            } else {
              console.warn( `Failed to copy: ${e_type}: element type is not supported` );
              target_body.appendParagraph(`\n>>>>> WARNING: Failed to copy ${e_type}: element type is not supported. <<<<<\n`);
              target_body.appendParagraph(`\n>>>>> WARNING: Refer to original source. <<<<<\n`);
            }
        }  
      } catch(e) {
        console.warn( `Failed to copy: ${e_type}: ${e}` );
        target_body.appendParagraph(`\n>>>>> WARNING: Failed to copy ${e_type}: "${e}". <<<<<\n`);
        target_body.appendParagraph(`\n>>>>> WARNING: Refer to original source. <<<<<\n`);
        // console.warn( source_body.getChild(i).getNumChildren() );
      }
    } 
  }

  /**
   * Prepares a document body for merging
   * @param {Object} source_body - body of existing Google Doc to copy from
   * Footnote elements are not handled well currently, so get their text and insert into parent element
   * (between double brackets [[]]) and then remove the elements from body.
   */
  replaceFootnotes( source_body ) {
    const searchType = DocumentApp.ElementType.FOOTNOTE;
    let searchResult = null;
    
    while (searchResult = source_body.findElement(searchType, searchResult)) {
      let footnote = searchResult.getElement().asFootnote();
      let footnote_parent = footnote.getParent();
      let footnote_index = footnote_parent.getChildIndex(footnote);
      let footnote_text = footnote.getFootnoteContents().getText();
      footnote_parent.insertText( footnote_index + 1, " [[ "+footnote_text+" ]] " );
      // footnote.removeFromParent();
    }
    
    searchResult = null;
    while (searchResult = source_body.findElement(searchType)) {
      let footnote = searchResult.getElement().asFootnote();
      footnote.removeFromParent();
    }
  }
}

/**
 * Enhanced version of JS Date class to interconvert with Google Date and TimeOfDay objects
 * see API docs at https://developers.google.com/classroom/reference/rest/v1/Date
 * 
 * @example
 *   const google_date = { year: 2023, month: 12, day: 2 };
 *   const google_time = { hours: 11, minutes: 0, seconds: 0 };
 *   const dt = DateTime.fromDateAndTime( google_date, google_time )
 *   dt.toISOString();  // 2023-12-02T11:00:00.000Z
 *   dt.asDate();  // {day=2.0, year=2023.0, month=12.0}
 *   dt.asTime();  // {minutes=0.0, hours=11.0, seconds=0.0, nanos=0.0}
 */
class DateTime extends Date {

  /**
   * create new DateTime from Google Date and TimeOfDay objects
   * @param {Object} date - Google Date
   * @param {Object} time - Google TimeOfDay
   * @return {DateTime} a new DateTime
   */
  static fromDateAndTime( date, time ) {
    return new DateTime( Date.UTC(
      date.year, date.month - 1, date.day,
      time.hours || 0, time.minutes || 0, time.seconds || 0
      // Google API may return a TimeOfDay object with missing properties
    ));
  }

  /**
   * create new DateTime from string representing date in ISO 8601
   * @param {String} isostring - date as "YYY-MM-DDTHH:mm:ss.sssZ"
   * @return {Object} a new DateTime
   */
  static fromISOString( isostring ) {
    return new DateTime( Date.parse( isostring ) );
  }

  /**
   * convert to Google Date object
   * @return {Object} as { year: , month: , day: }
   */
  asDate() {
    return { year: this.getUTCFullYear(), month: this.getUTCMonth() + 1, day: this.getUTCDate() };
  }

  /**
   * convert to Google TimeOfDay object
   * @return {Object} as { hours: , minutes: , seconds: , nanos: } in UTC
   */
  asTime() {
    return { 
      hours: this.getUTCHours(), minutes: this.getUTCMinutes(), 
      seconds: this.getUTCSeconds(), nanos: this.getUTCMilliseconds() * 1000000
    };
  }
}

/**
 * Base class to represent a resource returned by a Google API method.
 * Intended usage: derived classes should simply attach a static method
 * to create the specific type of resource; and then attach additional
 * methods appropriate for that resource.
 * 
 * @example
 *   static <createResource>() {
 *     let resource = <methodToGetResourceFromAPI>();
 *     return new this( resource );
 *   }
 */
class Resource {
  /** Assign each property from API return object to this */
  constructor( res ) { Object.assign( this, res ) }

  /**
   * convert to object containing only a subset of properties
   * @param {String[]} keys - Array of strings naming desired properties
   */
  subset( keys ) {
    return Object.assign( 
      ...keys.map( key => ( { [key] : this[key] } ) ) 
    );
  }
}

/**
 * Represent a Course resource from the Classroom API
 * 
 * @example
 *   const course1 = Course.getById( '523538288795' );
 *   const a1 = course1.getAssignments()[0];
 *   Logger.log ( a1.description );
 *   // "Create an informational product to help the audience make informed decisions about 
 *   // personal or public health regarding a health-related compound (medication, supplement, 
 *   // etc.). Attach or link all materials that are part of this product."
 *   
 * @example
 *   const course2 = Course.getById( 654295692323 );
 *   const assignments2 = course2.getAssignments();
 *   Logger.log ( assignments2.map( a => a.title ) );
 *   // [Week 08 reflection, Week 08 preparation, Week 07 reflection, Week 07 preparation, ... ]
 *   
 * @example
 *   const topic2 = course2.getTopicByName( "Journals" );
 *   Logger.log( topic2.topicId );
 */

class Course extends Resource {

  /**
   * create resource for existing course given its id
   * @param {string} course_id - unique id of a course
   * @return {Course} a new Course resource
   */
  static getById( course_id ) {
    let res = Classroom.Courses.get( course_id );
    return new this( res );
  }

  /**
   * fetch all courses that user has access to
   * @return {Course[]} Array of Course resources
   * 
   * affected by Settings.pageSize
   */
  static list(){
    let pageSize = Settings.pageSize;
    let res = Classroom.Courses.list( { pageSize: pageSize } );
    let pageToken = res.nextPageToken;
    let course_list = [ res.courses ];  // array of array, will flatten later
    while ( pageToken ) {
      res = Classroom.Courses.list( { pageSize: pageSize, pageToken: pageToken } );
      pageToken = res.nextPageToken;
      course_list.push( res.courses );
    }
    return course_list.flat().map( course => new this( course ) );
  }
  
  getAssignments()               { return Assignment.list( this.id ) }
  getAssignment( assignment_id ) { return Assignment.get( this.id, assignment_id ) }
  getMaterials()                 { return Material.list( this.id ) }
  getMaterial( material_id )     { return Material.get( this.id, material_id ) }
  getStudents()                  { return Student.list( this.id) }
  getStudent( user_id )          { return Student.get( this.id, user_id ) }
  getStudentByEmail( email )     { return Student.get( this.id, email ) }
  getTopics()                    { return Topic.list( this.id ) }
  getTopic( topic_id )           { return Topic.get( this.id, topic_id ) }
  
  /** 
   * search for named topic
   * @param {String} name
   * @return {Object} a new Topic resource, or undefined if not found in course
   */
  getTopicByName( name ) {
    let topic_list = this.getTopics();
    let matching_topics = topic_list ? topic_list.filter( t => t.name == name ) : [];
    if ( matching_topics.length > 0 ) { 
      return matching_topics[0];
    } else {
      return undefined;
    }
  }

  /** 
   * create a new topic in this Course, or gets existing one with this name
   * (Classroom API requires unique name per topic)
   * @param {String} name
   * @return {Object} a new Topic resource for the created or existing topic
   */
  createTopic( name ) {
    let existing_topic = this.getTopicByName( name );
    if ( existing_topic ) {
      return existing_topic;
    } else {
      return Topic.create( {name: name }, this.id );
    }
  }

  /** 
   * create a new Assignment in this Course, according to specification in:
   *   https://developers.google.com/classroom/reference/rest/v1/courses.courseWork
   * @param {Object} spec - field workType will be ignored and set to 'ASSIGNMENT'
   * @return {Object} a new Assignment resource for the created or existing topic
   */
  createAssignment( spec ) {
    const new_assignment = Assignment.create( spec, this.id );
    return new_assignment;
  }
}

/** 
 * Represent a Student resource from the Classroom API
 * ref: https://developers.google.com/classroom/reference/rest/v1/courses.students
 */
class Student extends Resource {

  /**
   * fetch all students enrolled in the course
   * @param {String} course_id
   * @return {Student[]} Array of Student resources
   * 
   * affected by Settings.pageSize
   */
  static list( course_id ) {
    let pageSize = Settings.pageSize;
    let res = Classroom.Courses.Students.list( course_id, { pageSize: pageSize } );
    let pageToken = res.nextPageToken;
    let student_list = [ res.students ];  // array of array, will flatten later
    while ( pageToken ) {
      res = Classroom.Courses.Students.list( course_id, { pageSize: pageSize, pageToken: pageToken } );
      pageToken = res.nextPageToken;
      student_list.push( res.students );
    }
    return student_list.flat().map( student => new this( student ) );
  }

  /**
   * create Resource for existing student given their user id
   * @param {String} course_id
   * @param {String} user_id
   * @return {Student} a new Student resource
   */
  static get( course_id, user_id ) {
    let res = Classroom.Courses.Students.get( course_id, user_id );
    return new Student( res );    
  }
  getEmail() { return this.profile.emailAddress }
  getFullName() { return this.profile.name.fullName }
  getWorkFolder() { return this.studentWorkFolder.alternateLink }
}

/** 
 * Represent a Topic resource from the Classroom API 
 * ref: https://developers.google.com/classroom/reference/rest/v1/courses.topics
 */
class Topic extends Resource {

  /**
   * fetch all topics in the given course
   * @param {String} course_id
   * @return {Topic[]} Array of Topic resources
   * 
   * affected by Settings.pageSize
   */
  static list( course_id ) {
    let pageSize = Settings.pageSize;
    let res = Classroom.Courses.Topics.list( course_id, { pageSize: pageSize } );
    let pageToken = res.nextPageToken;
    let topic_list = [ res.topic ];  // array of array, will flatten later
    while ( pageToken ) {
      res = Classroom.Courses.Topics.list( course_id, { pageSize: pageSize, pageToken: pageToken } );
      pageToken = res.nextPageToken;
      topic_list.push( res.topics );  
    }
    return topic_list.flat().map( topic => new this( topic ) );
  }
  
  /**
   * create Resource for existing topic given its id
   * @param {String} course_id
   * @param {String} topic_id
   * @return {Topic} a new Topic resource
   */
  static get( course_id, topic_id ) {
    let res = Classroom.Courses.Topics.get( course_id, topic_id );
    return new this( res );
  }

  /**
   * create a new topic in the given course
   * @param {String} course_id
   * @param {Object} spec - for a topic, e.g. {name: "Journals"}
   * @return {Topic} a new Topic resource
   * 
   * WARNING: other scripts may not have permission to patch or delete this topic
   */
  static create( spec, course_id ) {
    const res = Classroom.Courses.Topics.create( spec, course_id );
    return new this( res );
  }

  /**
   * @return {null}
   * 
   * WARNING: may not have permission to delete a topic that this project didn't create
   */
  remove() {
    Classroom.Courses.Topics.remove( this.courseId, this.topicId );
  }

  getName() { return this.name }
  getId() { return this.topicId }
}

/** 
 * Represent a Material resource from the Classroom API 
 * ref: https://developers.google.com/classroom/reference/rest/v1/courses.courseWorkMaterials
*/
class Material extends Resource {

  /**
   * fetch all materials in the given course
   * @param {String} course_id
   * @return {Material[]} Array of Material resources
   * 
   * affected by Settings.pageSize
   */
  static list( course_id ) {
    let pageSize = Settings.pageSize;
    let res = Classroom.Courses.CourseWorkMaterials.list( course_id, { 
      pageSize: pageSize, courseWorkStates: ['DRAFT', 'PUBLISHED'] 
    } );
    let pageToken = res.nextPageToken;
    let material_list = [ res.courseWorkMaterial ];  // array of array, will flatten later
    while ( pageToken ) {
      res = Classroom.Courses.CourseWorkMaterials.list( course_id, { 
        pageSize: pageSize, courseWorkStates: ['DRAFT', 'PUBLISHED'], pageToken: pageToken
      } );
      pageToken = res.nextPageToken;
      material_list.push( res.courseWorkMaterial );
    }
    return material_list.flat().map( material => new this( material ) );
  }

  /**
   * create Resource for existing material given its coursework id
   * @param {String} course_id
   * @param {String} coursework_id
   * @return {Material} a new Material resource
   */
  static get( course_id, coursework_id ) {
    let res = Classroom.Courses.CourseWorkMaterials.get( course_id, coursework_id );
    return new this( res );
  }

  /**
   * create a new material in the given course
   * @param {String} course_id
   * @param {Object} spec - for a material, e.g. {title: "Syllabus", ... }
   * @return {Material} a new Material resource
   * 
   * WARNING: other scripts may not have permission to patch or delete this material
   */
  static create( spec, course_id ) {
    const res = Classroom.Courses.CourseWorkMaterials( spec, course_id );
    return new this( res );
  }
}

/** 
 * Represent a CourseWork resource from the Classroom API 
 * ref: https://developers.google.com/classroom/reference/rest/v1/courses.courseWork
*/
class Assignment extends Resource {

  /**
   * fetch all coursework in the given course
   * @param {String} course_id
   * @return {Assignment[]} Array of Assignment resources
   * 
   * affected by Settings.pageSize
   */
  static list( course_id ) {
    let pageSize = Settings.pageSize;
    let res = Classroom.Courses.CourseWork.list( course_id, { 
      pageSize: pageSize, courseWorkStates: ['DRAFT', 'PUBLISHED'] 
    } );
    let pageToken = res.nextPageToken;
    let assignment_list = [ res.courseWork ];  // array of array, will flatten later
    while ( pageToken ) {
      res = Classroom.Courses.CourseWork.list( course_id, { 
        pageSize: pageSize, courseWorkStates: ['DRAFT', 'PUBLISHED'], pageToken: pageToken
      } );
      pageToken = res.nextPageToken;
      assignment_list.push( res.courseWork );
    }
    return assignment_list.flat().map( assignment => new this( assignment ) );
  }
  
  /**
   * create Resource for existing coursework given its coursework id
   * @param {String} course_id
   * @param {String} coursework_id
   * @return {Assignment} a new Assignment resource
   */
  static get( course_id, coursework_id ) {
    let res = Classroom.Courses.CourseWork.get( course_id, coursework_id );
    return new this( res );
  }

  /**
   * create a new coursework in the given course
   * @param {String} course_id
   * @param {Object} spec - for a CourseWork, e.g. { title: "Journal 01", ... }
   * @return {Assignment} a new Assignment resource
   * 
   * WARNING: other scripts may not have permission to patch or delete 
   * resources created by this script
   */
  static create( spec, course_id ) {
    const patched_spec = Object.assign( spec, { workType: 'ASSIGNMENT' } );
    const res = Classroom.Courses.CourseWork.create( patched_spec, course_id );
    return new this( res );
  }
  
  /**
   * @return {null}
   * 
   * WARNING: may not have permission to delete a topic that this project didn't create
   */
  remove() {
    Classroom.Courses.CourseWork.remove( this.courseId, this.id );
  }
  
  /** @return {DateTime} due date and time as a DateTime object */
  get due() {
    try {
      return DateTime.fromDateAndTime( this.dueDate, this.dueTime ); 
    } catch( err ) {
      return null;
    }
  }

  /**
   * @param {Date} date - Date or DateTime object, or empty string from sheet
   */
  set due( date ) {
    if( date > 0 ) {
      const datetime = new DateTime( date );
      this.dueDate = datetime.asDate();
      this.dueTime = datetime.asDate();
    } else {
      this.dueDate = null;
      this.dueTime = null;
    }
  }

  /** @return {DateTime} scheduled date and time as a DateTime object */
  get scheduled() {
    try {
      return DateTime.fromISOString( this.scheduledTime ); // DateTime object
    } catch( err ) {
      return null;
    }
  }

  /**
   * @param {Date} date - Date or DateTime object, or NaN
   */
  set scheduled( date ) {
    if( date > 0 ) {
      const datetime = new DateTime( date );
      this.scheduledTime = datetime.toISOString();
    } else {
      this.scheduledTime = null;
    }
  }

  /** @return {Submission[]} Array of Submission for this Assignment */
  getSubmissions() {
    return Submission.list( this.courseId, this.id );
  }

  getSubmission( submission_id ) {
    return Submission.get( this.courseId, this.id, submission_id );
  }
}

/** 
 * Represent a StudentSubmission resource from the Classroom API 
 * ref: https://developers.google.com/classroom/reference/rest/v1/courses.courseWork.studentSubmissions
 * ref: https://developers.google.com/classroom/reference/rest/v1/courses.courseWork.studentSubmissions#Attachment
 * ref: https://developers.google.com/classroom/reference/rest/v1/DriveFile
*/
class Submission extends Resource {
  
  /**
   * fetch all student submissions in the given course for given coursework
   * @param {String} course_id
   * @param {String} coursework_id
   * @return {Submission[]} Array of Submission resources
   * 
   * affected by Settings.pageSize
   */
  static list( course_id, coursework_id ) {
    let pageSize = Settings.pageSize;
    let res = Classroom.Courses.CourseWork.StudentSubmissions.list( 
      course_id, coursework_id, { pageSize: pageSize } 
    );
    let pageToken = res.nextPageToken;
    let submission_list = [ res.studentSubmissions ];  // array of array, will flatten later
    while ( pageToken ) {
      res = Classroom.Courses.CourseWork.StudentSubmissions.list( 
        course_id, coursework_id, { pageSize: pageSize, pageToken: pageToken,} 
      );
      pageToken = res.nextPageToken;
      submission_list.push( res.studentSubmissions );
    }
    return submission_list.flat().map( submission => new this( submission ) );
  }

  /**
   * create Resource for existing student submission given its submission id
   * @param {String} course_id
   * @param {String} coursework_id
   * @param {String} submission_id
   * @return {Submission} a new Submission resource
   */
  static get( course_id, coursework_id, submission_id ) {
    let res = Classroom.Courses.CourseWork.StudentSubmissions.get( course_id, coursework_id, submission_id );
    return new this( res );
  }

  /** @return {Object[]} Array of objects, each representing an attachment (empty if none) */
  get attachments() {
    return this.assignmentSubmission.attachments || [];
  }

  /** @return {Object[]} Array of objects, each representing drive file (empty if none) */
  get driveFiles() {
    let attachments = this.attachments.filter( a => a );
    let attachments_with_drivefiles = attachments.filter( a => "driveFile" in a );
    return attachments_with_drivefiles.map( a => a.driveFile );
  }
}


// TODO: If SheetTable encounters a column within DataRange with no label, create one
// TODO: updateAndPreserve should throw error if indexHeader isn't common to this and sheet

/**
 * Enhanced Object[] with helper methods to interact with Sheets
 * SheetTable enables conversion between API data and Sheet data.
 * It can read Range data from a sheet into a SheetTable (Array of Objects).
 * It can update a Sheet containing a listing of data from Resources.
 */
class SheetTable extends Array {

  /** @return {String[]} Array of strings naming column headers in order */
  get headers() { return Object.keys( this[0] ) }

  /**
   * Create new SheetTable with specified subset of properties
   * if any specified headers are not in the table, their values will be null
   * @param {String[]} headersArray - Array of strings that specify column headers
   * @return {SheetTable}
   */
  selectColumns( headersArray ) {
  // returns Table with only specified subset of properties
  // if any specified headers are not in the table, values will be null
    return this.map( row => {
      const entries = headersArray.map( header => {
        return [ header, row[header] ];
      });
      return Object.fromEntries( entries );
    });
  }

  /**
   * read values from sheet and interpret as a Table structure
   * @param {object} sheet
   * @param {number} [row_number] - row number containing headers, default 1
   * @return {SheetTable} imported data structured as a Table-like
   */
  static fromSheet( sheet, row_number = 1 ) {
    let rows = sheet.getDataRange().getValues();
    const filler = rows.splice( 0, row_number - 1 ); // intentional side effect: remove rows above header row
    const headers = rows.splice( 0, 1 )[0];          // intentional side effect: remove header row from rows
    let table = SheetTable.from( rows, row => {
      const entries = headers.map( (header,i) => {
        return [ header, row[i] ]
      } );
      return Object.fromEntries( entries );
    });
    return table;
  }

  /**
   * create a look-up table to associate each row with a unique id
   * indexHeader must match a value in this.headers
   * @param {String} indexHeader - name of column in sheet to use as index
   * @return {Object} of form { id1: row1, id2: row2 ... }
   */
  toLUT( indexHeader ) {
    const entries = this.map( row => {
      return [ row[indexHeader], row ]
    });
    return Object.fromEntries( entries );
  }

  /**
   * write Table data to sheet, but only for fields that exist in the sheet
   * (to avoid disrupting positional formatting in existing sheet)
   * @param {Sheet} sheet
   * @return {void}
   */
  update( sheet ) {
    const current = SheetTable.fromSheet( sheet );
    const updated = this.selectColumns( current.headers );
    updated.toSheet( sheet );
  }

  /**
   * write Table data to sheet, but only for fields that exist in the sheet; same as update(),
   * except preserve values in the sheet whose columns aren't in this SheetTable's headers
   * indexHeader must match a value in this.headers; values in this column must be unique
   * @param {Sheet} sheet
   * @param {String} indexHeader - name of column in sheet to use as index
   * @return {void}
   */
  updateAndPreserve( sheet, indexHeader ) {
    const current = SheetTable.fromSheet( sheet );
    const updated = this;
    
    const updated_LUT = updated.toLUT( indexHeader );
    const current_patched = current.map( row => {
      const index = row[ indexHeader ];
      // patch the row from current only if its index also exists in updated_LUT
      const patched_row = updated_LUT[ index ] || row;
      return Object.assign( row, patched_row );
    }).selectColumns( current.headers );

    const current_LUT = current_patched.toLUT( indexHeader );
    const updated_patched = updated.map( row => {
      const index = row[ indexHeader ];
      const patched_row = current_LUT[ index ] || row;
      return Object.assign( row, patched_row );
    });

    updated_patched.update( sheet );
  }

  /**
   * write headers to first row of sheet, then corresponding data to remaining rows
   * format first row as frozen
   * @param {Sheet} sheet - Sheet to receive data in this SheetTable
   * @return {void}
   */
  toSheet( sheet ) {
    if ( this.length < 1 ) throw 'SheetTable.toSheet: Table is empty';    
    const body = this.map( row => Object.values( row ) );
    const new_arr = [ this.headers ].concat(body);
    SheetTable.write( new_arr, sheet );
    SheetTable.clearEmpty( sheet );
    sheet.setFrozenRows(1);
  }

  /**
   * write array to sheet (overwrites entire data range)
   * @param {Array} arr - data to write
   * @param {Sheet} sheet - Sheet to receive data in arr
   * @return {void}
   */
  static write( arr, sheet ) {
    if ( arr.length < 1 ) throw 'Array is empty';
    const num_rows = arr.length;
    const num_cols = Math.max( arr[0].length, arr[1].length );
    // sheet.insertRowsAfter(1, num_rows);
    sheet.getDataRange().clearContent();
    let insert_range = sheet.getRange( 1, 1, num_rows, num_cols );
    insert_range.setValues( arr );
  }

  /**
   * detect minimum Range needed to contain existing data, and remove excess rows and columns
   * @param {Sheet} sheet
   * @return {void}
   */
  static clearEmpty( sheet ) {
    const last_row_pos = sheet.getLastRow();
    const target_rows = sheet.getMaxRows() - last_row_pos;
    if ( target_rows > 0 ) { sheet.deleteRows( last_row_pos + 1, target_rows ) }
    const last_col_pos = sheet.getLastColumn();
    const target_cols = sheet.getMaxColumns() - last_col_pos;
    if ( target_cols > 0 ) { sheet.deleteColumns( last_col_pos + 1, target_cols ) }
  }
}
