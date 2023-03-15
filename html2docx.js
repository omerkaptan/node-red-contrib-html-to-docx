const fs = require("fs");
const HTMLtoDOCX = require("html-to-docx");
const printDocx = async (html, filename, options) => {
  const filePath = filename || "";
  const htmlString = html || "testing";
  const optionsObj = options || {}
  const fileBuffer = await HTMLtoDOCX(htmlString, null, optionsObj);
  fs.writeFile(filePath, fileBuffer, (error) => {
    if (error) {
      return error;
    } else {
      // TODO: Başarı Durumunu return ederken Arayüzdeki Seçime Göre Ele Al <buffer|filepath>
    }
  });
};

module.exports = function (RED) {
  function html2Docx(config) {
    RED.nodes.createNode(this, config);
    const node = this;
    let {Output} = config
    node.on("input", async (msg, send, done) => {
      if (msg.hasOwnProperty("payload")) {
        try {
          const docx = await printDocx(msg.payload, msg.filename, msg.options);
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
