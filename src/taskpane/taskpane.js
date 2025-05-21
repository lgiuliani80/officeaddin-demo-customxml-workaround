/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
import { BlobReader, ZipReader, TextWriter } from '@zip.js/zip.js';

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("download-process").onclick = run;
  }
});

export async function run() {
  return Word.run(async (context) => {
    // Ottieni il documento come base64
        const doc = context.document;
        const file = Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 }, async (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const file = result.value;
                console.log('* File size: ' + file.size);
                console.log('* Slice size: ' + file.sliceSize);
                console.log('* Slice count: ' + file.sliceCount);
                let slices = [];
                let getSlice = (idx) => {
                    file.getSliceAsync(idx, (sliceResult) => {
                        if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
                            console.log('Slice ' + idx + ' size: ' + sliceResult.value.size);
                            slices.push(new Uint8Array(sliceResult.value.data));
                            if (slices.length < file.sliceCount) {
                                getSlice(idx + 1);
                                console.log('-> Next slice: ' + (idx + 1));
                            } else {
                                console.log('-> All slices received: ' + slices.length);
                                console.log(slices);
                                // Unisci i dati in un unico ArrayBuffer
                                const blob = new Blob(slices, { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
                                processDocx(blob);
                                file.closeAsync();
                            }
                        }
                    });
                };
                getSlice(0);
            }
        });
  });
}


async function processDocx(blob) {
    const output = document.getElementById('output');
    output.textContent = 'Analisi ZIP in corso...\r\n';
    output.textContent += 'Dimensione file: ' + blob.size + ' byte\r\n';
    // Converti il blob in base64
    const arrayBuffer = await blob.arrayBuffer();
    //const base64String = btoa(String.fromCharCode(...new Uint8Array(arrayBuffer)));
    //output.textContent += 'Base64: ' + base64String + '\r\n\r\n';
    try {
        const zipReader = new ZipReader(new BlobReader(blob));
        const entries = await zipReader.getEntries();
        for (let e of entries) {
            console.log('File: ' + e.filename + ' (' + e.uncompressedSize + ' byte)');
            if (e.filename.startsWith('customXml/') && e.filename.endsWith('.xml')) {
                console.log('-> Custom XML: ' + e.filename);
                const data = await e.getData(new TextWriter());
                console.log(data);
                output.textContent += 'Custom XML: ' + e.filename + ' - Length: ' + data.length + '\r\n';
                
                // Parsing dell'XML
                const parser = new DOMParser();
                const xmlDoc = parser.parseFromString(data, "application/xml");

                // Estrazione del namespace di root
                const rootNamespace = xmlDoc.documentElement.namespaceURI;
                output.textContent += '  namespace: ' + rootNamespace + '\r\n';
            }
        }
        
        //output.textContent += 'File nel docx:\n' + entries.map(e => e.filename).join('\r\n') + "\r\n";
        
        await zipReader.close();
    } catch (e) {
        output.textContent += 'Errore ZIP: ' + e + '\r\n';
    }
}
