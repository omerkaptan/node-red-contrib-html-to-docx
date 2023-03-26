# HTML to DOCX Node

The HTML to DOCX node converts HTML content to a DOCX file for node-red.

## Installation

To install the package, run the following command in your Node-RED user directory (typically `~/.node-red`):

```bash
npm install node-red-contrib-html-to-docx
```

## Usage

The HTML to DOCX node takes an HTML input and converts it to a DOCX output file. The output file name is specified using the `filename` property. In the example above, the output file is saved as "output.docx" and specified as a "file".

## Functionality

The function of this node is to convert the `msg.payload` content to a DOCX file and output it by saving it to the `msg.payload` property. Additionally, if the `output` property of the node is set to "Buffer", it returns a buffer instead of setting `msg.payload`.

## Properties

- **filename**: Specifies the name of the DOCX file to be saved.
- **output**: Specifies the output type. It can take the values "Buffer" or "file".

## Examples

The following example uses an `inject` node to send HTML content directly to the `html2docx` node and displays the output through a debug node.

```JSON
[{"id":"1eb5ec7d.4a9a8a","type":"inject","z":"73a31b55.9b0484","name":"","topic":"","payload":"<html><head><title>Test Page</title></head><body><h1>Hello, world!</h1></body></html>","payloadType":"str","repeat":"","crontab":"","once":false,"onceDelay":0.1,"x":190,"y":300,"wires":[["6e6672d9.75c7f"]]},{"id":"6e6672d9.75c7f","type":"html2docx","z":"73a31b55.9b0484","name":"","filename":"output.docx","output":"Buffer","x":430,"y":300,"wires":[["c1d6d956.6b9aa"]]},{"id":"c1d6d956.6b9aa","type":"debug","z":"73a31b55.9b0484","name":"","active":true,"tosidebar":true,"console":false,"tostatus":false,"complete":"false","statusVal":"","statusType":"auto","x":670,"y":300,"wires":[]}]
```

In this example, the `inject` node sends the HTML content directly to the `html2docx` node. The `html2docx` node converts the incoming HTML content to a DOCX file and sends the resulting buffer in the `msg.payload` property to the debug node.

## Error Handling

This node returns an error message when the HTML content cannot be converted. The error message can be retrieved using the `error` property added to the node's output.
