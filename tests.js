const TestSamples = {
  courses: {
    chem122: { id: '523538288795', assignment: '618713017845' },  // Chem 122, 10.5 preparation
    chem126: { id: '654295692323', assignment: '654297789443' },  // Chem 126, Week 07 preparation
    chemtas: { id: '545134506516', material: '1EWDUgX9ClqXejaghPpGmqdLCJxJACpx5JVUEZsAu5Mk' }
  }
}

function myTests() {
  // TestSheetTable.read_sheet();  // log headers and sample data from sheet 'courses'
  // TestSheetTable.write_sheet( 'test-to-sheet' );  // new sheet with 3x4 grid of data
  // TestSheetTable.sync_course_list( 'test-course-sync' );  // new sheet with updated course data
  // TestCourse.assignment_due();        // log due date and due time of Chem 126 Week 07 prep
  // TestCourse.assignments_list();      // log list of assignments in Chem 122, Chem 126, then 1 item from array
  // TestCourse.submissions_list();      // log first 5 submissions for Chem 122, 10.5 preparation
  // TestCourse.get_submission_first();  // log a submission in full for Chem 122, 10.5 preparation
  // TestCourse.students_list();         // log roster of Chem 122, email of first student
  // TestCourse.topics_list();           // log topics in Chem 122, then find topics "Journals" and "Chickens"
  // TestSubmissionMerge.merge_assignment_submissions( 'test submission merge' );  // new doc in MyDrive
  // Assignment.fromJournal()
  // TestCourse.create_journal( TestSamples.courses.chemtas.material );  // create assignment in chemTAs course, log spec
  // TestCourse.create_journal_from_sheet();  // create assignment in chemTAs course, log spec
  // Journal class to specify type of assignment
}

function cleanUpTests() {
  removeExistingSheetByName( 'test-to-sheet' );
  removeExistingSheetByName( 'test-course-sync' );
  const chem122 = Course.getById( TestSamples.courses.chem122.id );
  try{ chem122.getTopicByName( 'Chickens' ).remove(); } catch {}
  TestCourse.remove_journals();  // in {{TestSamples.courses.chemtas}}, created by create_journal()
}

function removeExistingSheetByName( name ) {
  const ss = SpreadsheetApp.getActive();
  let existing_sheet = ss.getSheetByName( name );
  if ( existing_sheet != null ) {
    ss.deleteSheet( existing_sheet );
  };
}

class TestDateTime extends DateTime {
  static test() {
    // instantiate DateTime
    let generic_date = new Date();
    let datetime = new DateTime( generic_date );
    Logger.log( generic_date.toISOString() );
    Logger.log( datetime.toISOString() );  // should be same as above
    
    // check for roundtripping
    let generic_date_to_ISO = generic_date.toISOString();
    let iso_to_datetime = DateTime.fromISOString( generic_date_to_ISO );
    Logger.log( iso_to_datetime );  // e.g. "Fri Mar 15 10:01:40 GMT-04:00 2024"

    // convert a Google Sheets Date to DateTime
    let test_date = { year: 2023, month: 12, day: 2 };
    let test_time = { hours: 11, minutes: 0, seconds: 0 };
    let dt = DateTime.fromDateAndTime( test_date, test_time )
    Logger.log( dt.toISOString() );  // 2023-12-02T11:00:00.000Z
    Logger.log( dt.asDate() );  // {day=2.0, year=2023.0, month=12.0}
    Logger.log( dt.asTime() );  // {minutes=0.0, hours=11.0, seconds=0.0, nanos=0.0}
  }  
}

class TestSheetTable extends SheetTable {

  static read_sheet() {
    const active = SpreadsheetApp.getActive();
    const sheet = active.getSheetByName( 'courses' ) || active.insertSheet( 'courses' );
    const table = SheetTable.fromSheet( sheet );
    Logger.log( table.headers );  // 	[id, section, courseState, alternateLink, name, include]
    Logger.log( table.selectColumns( ['id', 'section', 'courseState'] ) );
    // [ {courseState=ACTIVE, id=6.54295692323E11, section=Chem 126.01 Spring 2024}, 
    //   {courseState=ACTIVE, id=5.23836380627E11, section=Chem 123.06 Fall 2023}, 
    //   {courseState=ACTIVE, id=5.23747147687E11, section=Chem 123.01 Fall 2023}, ... ]
  }

  static write_sheet( test_sheet_name ) {
    removeExistingSheetByName( test_sheet_name );

    const table = SheetTable.from( [ 
      { id: 1, name: "A", include: true },
      { id: 2, name: "B", include: false },
      { id: 3, name: "C", include: true },
    ]);

    const sheet = SpreadsheetApp.getActive().insertSheet( test_sheet_name );
    table.toSheet( sheet );
    // New sheet with 3 x 4 grid of cells:
    // | id | name | include |
    // | 1  | A    | TRUE    |
    // | 2  | B    | FALSE   |
    // | 3  | C    | TRUE    |
  }

  static sync_course_list( test_sheet_name ) {
    const ss = SpreadsheetApp.getActive();
    removeExistingSheetByName( test_sheet_name );
    let course_sheet = ss.getSheetByName('courses');
    let test_sheet = course_sheet.copyTo( ss ).setName( test_sheet_name );
    
    const course_list = Course.list();
    const data = SheetTable.from( course_list );
    data.updateAndPreserve( test_sheet, 'id' );
  }

  static update() {
    const course_list = Course.list();
    let data = SheetTable.from( course_list );
    let sheet = SpreadsheetApp.getActive().getSheetByName('test');
    data.update( sheet );
  }

  static update_and_preserve_courses() {
    const course_list = Course.list();
    let data = SheetTable.from( course_list );
    let sheet = SpreadsheetApp.getActive().getSheetByName('test');
    data.updateAndPreserve( sheet, 'id' );
  }

  static update_and_preserve_assignments() {
    const assignment_list = Course.getById( TestSamples.courses.chem122.id ).getAssignments();
    let data = SheetTable.from( assignment_list );
    let sheet = SpreadsheetApp.getActive().getSheetByName('test');
    data.updateAndPreserve( sheet, 'id' );
  }

  static update_and_preserve_submissions() {
    const course = Course.getById( TestSamples.courses.chem122.id );
    const submission_list = course.getAssignment( TestSamples.courses.chem122.assignment ).getSubmissions();
    let data = SheetTable.from( submission_list );
    let sheet = SpreadsheetApp.getActive().getSheetByName('test');
    data.updateAndPreserve( sheet, 'id' );
  }
}

class TestMerge extends MergeDoc {

  static merge_generic() {
    let test_id = '198DzO6fBnk5VOlJWVou5wRrc-t8cozXTrBXfcB-Pf2I';  // Journal reflection document
    let merge_doc = new MergeDoc( 'test generic merge' );
    let source_doc = DocumentApp.openById( test_id );
    merge_doc.addTitle( 'Title 1' );
    merge_doc.addSubtitle( 'Subtitle 1' );
    merge_doc.addDocument( source_doc );
    merge_doc.addTitle( 'Title 1', 'www.google.com' );
    merge_doc.addSubtitle( 'Subtitle 1', 'www.google.com' );
    merge_doc.addDocument( source_doc );
    Logger.log( `New document at ${merge_doc.doc.getUrl()}` );
  }

  static check_doc_structure() {
    // displays type of each element in doc
    let merge_doc = new MergeDoc( 'test generic merge' );
    let merge_body = merge_doc.doc.getBody();
    let doc = DocumentApp.openById( '1I15Ie0vrA4JTflUNsqQfg7YvPBuFpken9DMEBN6fAFM' );
    let body = doc.getBody();
    const num_elements = body.getNumChildren();
    for (let i=0; i < num_elements; i++ ) {
      let e = body.getChild(i).copy();
      let e_type = e.getType();
      try {
        merge_body.appendParagraph( e );
        Logger.log(`${i}: ${e_type.toString()}: ${e}`);
      } catch( error) {
        Logger.log(`${i}: ${e_type.toString()}: ${error}`);
      }
    }
    Logger.log( `New document at ${merge_doc.doc.getUrl()}` );
  }

  static inspect_problem_paragraphs() {
    let test_id = '1I15Ie0vrA4JTflUNsqQfg7YvPBuFpken9DMEBN6fAFM';
    // document by Ashlyn Widmer that contains pasted floating images
    let merge_doc = new MergeDoc( 'test copy document with positioned images' );
    let source_doc = DocumentApp.openById( test_id );
    merge_doc.addDocument( source_doc );
    Logger.log( `New document at ${ merge_doc.doc.getUrl() }` );    
  }
}

class TestSubmissionMerge extends Course {
  static merge_assignment_submissions( merge_doc_name) {
    let merge_doc = new MergeDoc( merge_doc_name );
    let course = Course.getById( TestSamples.courses.chem126.id );
    let assignment = course.getAssignment( TestSamples.courses.chem126.assignment );  // Chem 126, Week 07 prep
    let submissions = assignment.getSubmissions();
    let index = Object.fromEntries( course.getStudents().map( s => [ s.userId, s.getEmail() ] ) );
    submissions.sort( (a,b) => {
      return index[a.userId] > index[b.userId] ? 1 : index[a.userId] < index[b.userId] ? -1 :0
    });
    submissions.forEach( submission => {
      submission.driveFiles.forEach( drivefile => {
        let owner = course.getStudent( submission.userId ).getEmail();
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
    return merge_doc.doc.getUrl();
  }
}

class TestCourse extends Course {

  static create_journal(  ) {
    Logger.log( "~~~  Running TestCourse.create_journal()  ~~~~~~~~~~~~~~~~~~~~~~~~~");
    
    const course = Course.getById( TestSamples.courses.chemtas.id );
    const topic = course.getTopicByName( 'Journals' );
    const materials_id = TestSamples.courses.chemtas.material;
    const dueJS = new DateTime( Date.UTC( 2025, 11, 1, 8+4, 30, 0 ) );  // 2025-Dec-01 at 0830 EST
    const schJS = new DateTime( Date.UTC( 2025, 10, 1, 8+4, 30, 0 ) );  // 2025-Nov-01 at 0830 EST
    const spec = { 
      topicId: topic.topicId, title: "Journal A", maxPoints: 55, state: 'DRAFT', 
      materials: [{ driveFile: { driveFile: { id: materials_id }, shareMode: 'STUDENT_COPY'} }],
      dueDate: dueJS.asDate(), dueTime: dueJS.asTime(), scheduledTime: schJS.toISOString(),
      description: "Hi!",
    };
    const new_assignment = course.createAssignment ( spec );
    Logger.log( JSON.stringify( new_assignment, null, 2 ) );
  }

  static create_journal_from_sheet() {
    Logger.log( "~~~  Running TestCourse.create_journal()  ~~~~~~~~~~~~~~~~~~~~~~~~~");

    const active = SpreadsheetApp.getActive();
    const sheet = active.getSheetByName( 'batch' );
    const table = SheetTable.fromSheet( sheet );
    const row = table[0];
    const journal = Journal.fromObject( row );
    const course = Course.getById( row.courseId );
    const new_assignment = course.createAssignment ( journal );
    Logger.log( JSON.stringify( new_assignment, null, 2 ) );
    return new_assignment;    
  }

  static remove_journals() {
    Logger.log( "~~~  Running TestCourse.remove_journals()  ~~~~~~~~~~~~~~~~~~~~~~~~~");
    
    const course = Course.getById( TestSamples.courses.chemtas.id );
    const topic = course.getTopicByName( 'Journals' );
    const journals = course.getAssignments().filter( journal => journal.topicId == topic.topicId )
    journals.forEach( journal => {
      journal.remove();
    });
  }

  static submissions_list() {
    Logger.log( "~~~  Running TestCourse.submissions_list()  ~~~~~~~~~~~~~~~~~~~~~~~~~");

    let course = Course.getById( TestSamples.courses.chem122.id );
    let assignment = course.getAssignment( TestSamples.courses.chem122.assignment ); 
    let submissions = assignment.getSubmissions();
    const submissions_description = submissions.map( sub => {
      return `${sub.id}: ${sub.driveFiles[0].id}, ${sub.driveFiles[0].title}`;
    })
    Logger.log( JSON.stringify( submissions_description.slice(0,5), null, 2 ) );
    // [ Cg4Ikavf7b4OEPWKrfGAEg: 1i7M1JIu7BM_iopD3fa-gXfYeiQbGzBS4fvhKMMUgwHM, Brittany <redacted> - Journal preparation,
    // ... ]
  }

  static get_submission_first() {
    Logger.log( "~~~  Running TestCourse.get_submission_first()  ~~~~~~~~~~~~~~~~~~~~~~~~~");

    const course = Course.getById( TestSamples.courses.chem122.id );
    const assignment = course.getAssignment( TestSamples.courses.chem122.assignment ); 
    const submissions = assignment.getSubmissions();
    const submission = submissions[0]
    Logger.log( JSON.stringify( submission, null, 2 ) );
    // { "state": "RETURNED",
    //   "submissionHistory": [
    //   ...
  }

  static assignment_due() {
    Logger.log( "~~~  Running TestCourse.assignment_due()  ~~~~~~~~~~~~~~~~~~~~~~~~~");

    let course = Course.getById( TestSamples.courses.chem126.id );
    let assignment = course.getAssignment( TestSamples.courses.chem126.assignment );  
    Logger.log( `Due DT: ${ assignment.dueDate }` );  // Due DT: {"month":2,"day":25,"year":2024}
    Logger.log( `Due DT: ${ assignment.dueTime }` );  // Due DT: {"hours":21}
    Logger.log( `Due DT: ${ assignment.due }` );      // Due DT: Sun Feb 25 2024 16:00:00 GMT-0500 (Eastern Standard Time)
  }

  static assignments_list() {
    Logger.log( "~~~  Running TestCourse.assignments_list()  ~~~~~~~~~~~~~~~~~~~~~~~~~");

    let course1 = Course.getById( TestSamples.courses.chem122.id );
    let course2 = Course.getById( 654295692323 );
    let a1 = course1.getAssignments()[0];
    Logger.log ( a1.description );
    // "Create an informational product to help the audience make informed decisions about 
    // personal or public health regarding a health-related compound (medication, supplement, 
    // etc.). Attach or link all materials that are part of this product."
    Logger.log ( course2.getAssignments().map( a => a.title ) );
    // [Week 08 reflection, Week 08 preparation, Week 07 reflection, Week 07 preparation, ... ]
    Logger.log ( JSON.stringify( course1.getAssignment( a1.id ) ) );
    // {creatorUserId=106367787409346626774, ..., title=Project, ... , maxPoints=10.0}
  }

  static students_list() {
    Logger.log( "~~~  Running TestCourse.students_list()  ~~~~~~~~~~~~~~~~~~~~~~~~~");

    let course = Course.getById( TestSamples.courses.chem122.id );
    let students = course.getStudents();
    const student_list = students.map( student => {
      return `${student.userId}: ${student.getEmail()}, ${student.getFullName()}`;
    });
    Logger.log( JSON.stringify( student_list ) );
    Logger.log( course.getStudent( students[0].userId ).getEmail() );
  }

  static topics_list() {
    Logger.log( "~~~  Running TestCourse.topics_list()  ~~~~~~~~~~~~~~~~~~~~~~~~~");

    let course = Course.getById( TestSamples.courses.chem122.id );  // Chem 122
    let topics = course.getTopics();
    let topics_list = topics.map( topic => `${topic.topicId}: ${topic.name}` );
    Logger.log( topics_list );  
    // 	[618722308662: Portfolio, 618722364401: Journals]

    const names = [ "Journals", "Chickens" ];
    names.forEach( name => {
      let topic = course.getTopicByName( name );
      Logger.log( `${ name }: ${JSON.stringify( topic )}`)
      course.createTopic( name );
    });
    // Journals: {"courseId":"523538288795", ..., "name":"Journals"}
    // Chickens: undefined
  }

  static instantiate() {
    // instantiate courses and get assignments
    let course1 = Course.getById( '523538288795' );
    let course2 = Course.getById( 654295692323 );
    let a1 = course1.getAssignments()[0];
    Logger.log ( a1.description );
    // "Create an informational product to help the audience make informed decisions about 
    // personal or public health regarding a health-related compound (medication, supplement, 
    // etc.). Attach or link all materials that are part of this product."
    Logger.log ( course2.getAssignments().map( a => a.title ) );
    // [Week 08 reflection, Week 08 preparation, Week 07 reflection, Week 07 preparation, ... ]
    Logger.log ( course1.getAssignment( a1.id ) )
    // {creatorUserId=106367787409346626774, ..., title=Project, ... , maxPoints=10.0}
  }

}
