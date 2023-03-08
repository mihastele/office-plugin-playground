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
