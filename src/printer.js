const net = require("net");

function printZebra(ip, zpl) {
  return new Promise((resolve, reject) => {
    const client = new net.Socket();

    client.connect(9100, ip, () => {
      client.write(zpl);
      client.end();
    });

    client.on("close", () => resolve("Impreso correctamente"));
    client.on("error", (err) => reject(err));
  });
}

module.exports = { printZebra };