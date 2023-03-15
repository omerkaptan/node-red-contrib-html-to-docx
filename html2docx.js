const fs = require("fs");
const { promisify } = require("util");
const HTMLtoDOCX = require("html-to-docx");
const writeFileAsync = promisify(fs.writeFile);
const printDocx = async (html, filename, options) => {
  const filePath = filename || "";
  const htmlString = html || "testing";
  const optionsObj = options || {};
  const fileBuffer = await HTMLtoDOCX(htmlString, null, optionsObj)
  writeFileAsync(filePath, fileBuffer);
  return fileBuffer;
};

module.exports = function (RED) {
  function html2Docx(config) {
    RED.nodes.createNode(this, config);
    const node = this;
    let { output } = config;
    node.on("input", async (msg, send, done) => {
      if (msg.hasOwnProperty("payload")) {
        try {
          let docx = await printDocx(msg.payload, msg.filename, msg.options)
            .then((fileBuffer) => {
              console.log("File buffer:", fileBuffer);
              msg.payload = docx || "1";
              send(msg);
              done();
            })
            .catch((error) => {
              done(error);
            });
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
