const fs = require("fs");
const { promisify } = require("util");
const HTMLtoDOCX = require("html-to-docx");

const writeFileAsync = promisify(fs.writeFile);

async function printDocx(html = "", filename = "", options = {}) {
  const fileBuffer = await HTMLtoDOCX(html, null, options || {});
  if (options.output === "Buffer") return fileBuffer;
  await writeFileAsync(filename, fileBuffer);
  return filename;
}

module.exports = function (RED) {
  function html2Docx(config) {
    RED.nodes.createNode(this, config);
    const node = this;
    let output = config.output;
    node.on("input", async (msg, send, done) => {
      if (msg?.payload) {
        try {
          const res = await printDocx(msg.payload, msg.filename, {
            ...output,
            ...msg.options,
          });
          msg.payload = res;
          send(msg);
          done();
        } catch (error) {
          done(error);
        }
      } else {
        send(msg);
        done();
      }

      if (done) done();
    });
  }
  RED.nodes.registerType("html2docx", html2Docx);
};
