Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("italic").onclick = italic;
    document.getElementById("bold").onclick = bold;
    document.getElementById("repeat").onclick = repeat;
    document.getElementById("fs_inc").onclick = fontSizeIncrease;
    document.getElementById("fs_dec").onclick = fontSizeDecrease;
    document.getElementById("random_string").onclick = insertRandomString;
    document.getElementById("custom_input-btn").onclick = insertCustomString;
    document.getElementById("left_align").onclick = textAlignLeft;
    document.getElementById("center_align").onclick = textAlignCenter;
    document.getElementById("right_align").onclick = textAlignRight;
  }
});

function bold() {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();

    selection.load("font");
    await context.sync();

    try {
      if (selection.font.bold) {
        selection.font.bold = false;
      } else {
        selection.font.bold = true;
      }
    } catch (error) {
      console.error(error.message);
    }
  });
}

function italic() {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("font");

    await context.sync();

    try {
      if (selection.font.italic) {
        selection.font.italic = false;
      } else {
        selection.font.italic = true;
      }
    } catch (error) {
      console.error(error.message);
    }
  });
}

function repeat() {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();

    try {
      context.document.body.insertText(
        selection.text.replaceAll("\r", ""),
        Word.InsertLocation.end
      );
    } catch (error) {
      console.error(error.message);
    }
  });
}

function fontSizeIncrease() {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("font");
    await context.sync();

    try {
      let currentSize = selection.font.size;
      let newSize = currentSize + 1;
      selection.font.size = newSize;
    } catch (error) {
      console.error(error.message);
    }
  });
}

function fontSizeDecrease() {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("font");
    await context.sync();

    try {
      let currentSize = selection.font.size;
      let newSize = currentSize - 1;
      selection.font.size = newSize;
    } catch (error) {
      console.error(error.message);
    }
  });
}

////////////////////////////////////
////// WITHOUT TEXT SELECTION //////
////////////////////////////////////

function insertRandomString() {
  return Word.run(async (context) => {
    try {
      context.document.body.insertText(
        "This is a random string",
        Word.InsertLocation.end
      );
    } catch (error) {
      console.error(error.message);
    }
  });
}

function insertCustomString() {
  return Word.run(async (context) => {
    try {
      const input = document.getElementById("custom_input");
      const text = input.value;
      if (text) {
        context.document.body.insertText(text, Word.InsertLocation.end);
      } else {
        throw new Error("Please type something in the input box.");
      }
    } catch (error) {
      console.error(error.message);
    }
  });
}

////////////////////////////////////
///////// TEXT ALIGNMENT ///////////
////////////////////////////////////

function textAlignLeft() {
  return Word.run(async (context) => {
    try {
      const body = context.document.body;
      body.load("paragraphs");
      await context.sync();

      const paragraphs = body.paragraphs.items;
      paragraphs.forEach((para) => {
        para.alignment = "Left";
      });
    } catch (error) {
      console.error(error.message);
    }
  });
}

function textAlignCenter() {
  return Word.run(async (context) => {
    try {
      const body = context.document.body;
      body.load("paragraphs");
      await context.sync();

      const paragraphs = body.paragraphs.items;
      paragraphs.forEach((para) => {
        para.alignment = "Centered";
      });
    } catch (error) {
      console.error(error.message);
    }
  });
}

function textAlignRight() {
  return Word.run(async (context) => {
    try {
      const body = context.document.body;
      body.load("paragraphs");
      await context.sync();

      const paragraphs = body.paragraphs.items;
      paragraphs.forEach((para) => {
        para.alignment = "Right";
      });
    } catch (error) {
      console.error(error.message);
    }
  });
}
