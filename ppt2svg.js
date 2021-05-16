const { PDFNet } = require('@pdftron/pdfnet-node');
const { join } = require('path');

const cwd = process.cwd();

const outputPath = join(cwd, 'output/simple') + "/";

const inputPath = cwd + '/';

const ResourcePath = cwd + '/Resource';


const pptToPdf = async (inputFilename, outputFilename) => {
    const pdfdoc = await PDFNet.Convert.officeToPdfWithPath(inputPath + inputFilename);
    await pdfdoc.save(outputPath + outputFilename, PDFNet.SDFDoc.SaveOptions.e_linearized);
    // And we're done!
    console.log('Saved ' + outputFilename);
}


const pptToSvg = async () => {
    PDFNet.addResourceSearchPath(ResourcePath);
    const pdfdoc = await PDFNet.Convert.officeToPdfWithPath(inputPath + 'p1.pptx');
   await PDFNet.Convert.docToSvg(pdfdoc, 'output/svg4/p1.svg')
}










PDFNet.runWithCleanup(pptToSvg).catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
  }).then(function () { PDFNet.shutdown(); });