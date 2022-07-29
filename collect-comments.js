var INCLUDE_AUTHOR = false;

function myFunction() {
  // Creates section at the top that lists each comment's context and content
  var doc = DocumentApp.getActiveDocument();
  var doc_id = doc.getId();
  var body = doc.getBody();
  var summary = body.insertParagraph( 0, "Comments" ).setHeading( DocumentApp.ParagraphHeading.NORMAL );
  
  var comments = Drive.Comments.list( doc_id ).items;
  comments.reverse().map( comment => {
    let cited_text = comment.context.value;
    let annotation = INCLUDE_AUTHOR ? 
      `${comment.author.displayName}: ${comment.content}` :
      `${comment.content}`;
    body.insertListItem( 1, cited_text ).setNestingLevel(0);
    body.insertListItem( 2, annotation ).setNestingLevel(1);
    comment.replies.reverse().map( reply =>{
      body.insertListItem( 3, reply.content ).setNestingLevel(1);
    })
    Logger.log( comment );
  })
  summary.setHeading( DocumentApp.ParagraphHeading.HEADING1 );
}
