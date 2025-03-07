# Manage comments

This sample shows how to perform basic comments operations, including insert, reply, get, edit, resolve, and delete.

This sample shows basic operations using comments.

```typescript
async function insertComment() {
  // Sets a comment on the selected content.
  await Word.run(async (context) => {
    const text = $("#comment-text")
      .val()
      .toString();
    const comment: Word.Comment = context.document.getSelection().insertComment(text);

    // Load object to log in the console.
    comment.load();
    await context.sync();

    console.log("Comment inserted:", comment);
  });
}

async function editFirstCommentInSelection() {
  // Edits the first active comment in the selected content.
  await Word.run(async (context) => {
    const text = $("#edit-comment-text")
      .val()
      .toString();
    const comments: Word.CommentCollection = context.document.getSelection().getComments();
    comments.load("items");
    await context.sync();

    const firstActiveComment: Word.Comment = comments.items.find((item) => item.resolved !== true);
    if (!firstActiveComment) {
      console.warn("No active comment was found in the selection, so couldn't edit.");
      return;
    }

    firstActiveComment.content = text;

    // Load object to log in the console.
    firstActiveComment.load();
    await context.sync();

    console.log("Comment content changed:", firstActiveComment);
  });
}

async function replyToFirstActiveCommentInSelection() {
  // Replies to the first active comment in the selected content.
  await Word.run(async (context) => {
    const text = $("#reply-text")
      .val()
      .toString();
    const comments: Word.CommentCollection = context.document.getSelection().getComments();
    comments.load("items");
    await context.sync();

    const firstActiveComment: Word.Comment = comments.items.find((item) => item.resolved !== true);
    if (firstActiveComment) {
      const reply: Word.CommentReply = firstActiveComment.reply(text);
      console.log("Reply added.");
    } else {
      console.warn("No active comment was found in the selection, so couldn't reply.");
    }
  });
}

async function toggleResolvedStatusOfFirstCommentInSelection() {
  // Toggles Resolved status of the first comment in the selected content.
  await Word.run(async (context) => {
    const comment: Word.Comment = context.document
      .getSelection()
      .getComments()
      .getFirstOrNullObject();
    comment.load("resolved");
    await context.sync();

    if (comment.isNullObject) {
      console.warn("No comments in the selection, so nothing to toggle.");
      return;
    }

    // Toggle resolved status.
    // If the comment is active, set as resolved.
    // If it's resolved, set resolved to false.
    const resolvedBefore = comment.resolved;
    console.log(`Comment Resolved status (before): ${resolvedBefore}`);
    comment.resolved = !resolvedBefore;
    comment.load("resolved");
    await context.sync();

    console.log(`Comment Resolved status (after): ${comment.resolved}`);
  });
}

async function getFirstCommentRangeInSelection() {
  // Gets the range of the first comment in the selected content.
  await Word.run(async (context) => {
    const comment: Word.Comment = context.document.getSelection().getComments().getFirstOrNullObject();
    comment.load("contentRange");
    const range: Word.Range = comment.getRange();
    range.load("text");
    await context.sync();

    if (comment.isNullObject) {
      console.warn("No comments in the selection, so no range to get.");
      return;
    }

    console.log(`Comment location: ${range.text}`);
    const contentRange: Word.CommentContentRange = comment.contentRange;
    console.log("Comment content range:", contentRange);
  });
}

async function getCommentsInSelection() {
  // Gets the comments in the selected content.
  await Word.run(async (context) => {
    const comments: Word.CommentCollection = context.document.getSelection().getComments();

    // Load objects to log in the console.
    comments.load();
    await context.sync();

    console.log("Comments:", comments);
  });
}

async function getRepliesToFirstCommentInSelection() {
  // Gets the replies to the first comment in the selected content.
  await Word.run(async (context) => {
    const comment: Word.Comment = context.document.getSelection().getComments().getFirstOrNullObject();
    comment.load("replies");
    await context.sync();

    if (comment.isNullObject) {
      console.warn("No comments in the selection, so no replies to get.");
      return;
    }

    const replies: Word.CommentReplyCollection = comment.replies;
    console.log("Replies to the first comment:", replies);
  });
}

async function deleteFirstCommentInSelection() {
  // Deletes the first comment in the selected content.
  await Word.run(async (context) => {
    const comment: Word.Comment = context.document.getSelection().getComments().getFirstOrNullObject();
    comment.delete();
    await context.sync();

    if (comment.isNullObject) {
      console.warn("No comments in the selection, so nothing to delete.");
      return;
    }

    console.log("Comment deleted.");
  });
}

async function getComments() {
  // Gets the comments in the document body.
  await Word.run(async (context) => {
    const comments: Word.CommentCollection = context.document.body.getComments();

    // Load objects to log in the console.
    comments.load();
    await context.sync();

    console.log("All comments:", comments);
  });
}

async function setup() {
  await Word.run(async (context) => {
    const body: Word.Body = context.document.body;
    body.clear();
    body.insertParagraph(
      "Do you want to create a solution that extends the functionality of Word? You can use the Office Add-ins platform to extend Word clients running on the web, on a Windows desktop, or on a Mac.",
      "Start"
    );
    body.paragraphs
      .getLast()
      .insertText(
        "Use add-in commands to extend the Word UI and launch task panes that run JavaScript that interacts with the content in a Word document. Any code that you can run in a browser can run in a Word add-in. Add-ins that interact with content in a Word document create requests to act on Word objects and synchronize object state.",
        "Replace"
      );
  });
}
// Default helper for invoking an action and handling errors.
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}
```

