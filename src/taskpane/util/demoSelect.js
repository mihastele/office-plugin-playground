export function getSelectedText() {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();
    return selection.text;
  });
}

export function appendTextToHTML(text) {
  document.getElementById("append-section").innerHTML += text;
}

function log(message) {
  console.log(message);
}

export function listenToSelectionChange() {
  return Word.run(async (context) => {
    // Register the event handler.
    //  context.document.addHandlerAsync(Word.EventType.documentSelectionChanged, log);

    console.log(context.binding);
    console.log("HIII");

    // console.log(Office);
    // console.log(Office.context);
    // console.log(Office.context.ui);
    // console.log(Office.context.ui.addHandlerAsync);
    // console.log(context);
    // console.log(context.document);

    // Office.context.document.addHandlerAsync(Word.EventType.documentSelectionChanged, log);

    /* Office.context.document.bindings.addFromSelectionAsync(
      Office.BindingType.Text,
      { id: "selectBind" },
      function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
          write("Action failed. Error: " + asyncResult.error.message);
        } else {
          write("Added new binding with type: " + asyncResult.value.type + " and id: " + asyncResult.value.id);
        }
      }
    );*/

    console.log(Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, log));

    console.log(Office.context.document.bindings.addFromNamedItemAsync());

    // Function that writes to a div with id='message' on the page.
    function write(message) {
      // document.getElementById("message").innerText += message;
      console.log(message);
    }

    await context.sync();
  });
}

/*
export function listenToSelectionChange() {
  return Word.run(async (context) => {
    // Register the event handler.
    context.document.addHandlerAsync(Word.EventType.documentSelectionChanged, trackMessage);
    await context.sync();
  });
}
*/
