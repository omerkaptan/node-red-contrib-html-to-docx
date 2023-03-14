const fs = require("fs");
const HTMLtoDOCX = require("html-to-docx");
const filePath = "C:\\Users\\omer.kaptan\\test.docx";

const printDocx = async (html, options) => {
  const htmlString = html || "testing";
  const fileBuffer = await HTMLtoDOCX(htmlString, null, {});

  fs.writeFile(filePath, fileBuffer, (error) => {
    if (error) {
      return error;
    }
  });
};

module.exports = function (RED) {
  function html2Docx(config) {
    RED.nodes.createNode(this, config);
    var node = this;
    node.on("input", async (msg, send, done) => {
      if (msg.hasOwnProperty("payload")) {
        try {
          const docx = await printDocx(msg.payload, this.options);
          msg.payload = docx || "1";
          send(msg);
          done();
        } catch (error) {
          done(error);
        }
      } else {
        // If no payload just pass it on.
        send(msg);
        done();
      }

      // Out

      // Once finished, call 'done'.
      // This call is wrapped in a check that 'done' exists
      // so the node will work in earlier versions of Node-RED (<1.0)
      if (done) {
        done();
      }
    });
  }
  RED.nodes.registerType("html2docx", html2Docx);
};
