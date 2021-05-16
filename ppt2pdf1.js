const { PDFNet } = require('@pdftron/pdfnet-node');


const runOfficeToPDF = async () => {
    const inputPath = './TestFiles/';
    const outputPath = inputPath + 'Output/';

    const simpleDocxConvert = async (inputFilename, outputFilename) => {
        const pdfdoc = await PDFNet.Convert.officeToPdfWithPath(inputPath + inputFilename);
        await pdfdoc.save(outputPath + outputFilename, PDFNet.SDFDoc.SaveOptions.e_linearized);
        console.log('Saved ' + outputFilename);
    }

    const pptToSvg = async (inputFilename, outputFilename) => {
        const pdfdoc = await PDFNet.Convert.officeToPdfWithPath(inputPath + inputFilename);
        console.log('pdfdoc ==> ', pdfdoc);
        try {
            await PDFNet.Convert.docToSvg(pdfdoc, outputPath + outputFilename)
        } catch (error) {
            console.log('error ===> ', error);
        }
    }

    const flexibleDocxConvert = async (inputFilename, outputFilename) => {
        const pdfdoc = await PDFNet.PDFDoc.create();
        pdfdoc.initSecurityHandler();
        const options = new PDFNet.Convert.OfficeToPDFOptions();
        options.setSmartSubstitutionPluginPath(inputPath);
        const conversion = await PDFNet.Convert.streamingPdfConversionWithPdfAndPath(pdfdoc, inputPath + inputFilename, options);

        while (await conversion.getConversionStatus() === PDFNet.DocumentConversion.Result.e_Incomplete) {
            await conversion.convertNextPage();
        }

        if (await conversion.getConversionStatus() === PDFNet.DocumentConversion.Result.e_Success) {
            const num_warnings = await conversion.getNumWarnings();
            for (let i = 0; i < num_warnings; ++i) {
                console.log('Conversion Warning: ' + await conversion.getWarningString(i));
            }

            // save the result
            await pdfdoc.save(outputPath + outputFilename, PDFNet.SDFDoc.SaveOptions.e_linearized);
            // done
            console.log('Saved ' + outputFilename);
        } else {
            console.log('Encountered an error during conversion: '
                + await conversion.getErrorString());
        }
    }

    const main = async () => {
        PDFNet.addResourceSearchPath('./Resources');

        try {
            await simpleDocxConvert('mysql.pptx', 'mysql.pdf');
            await pptToSvg('mysql.pptx', '/sv5/mysql.svg');
            await pptToSvg('p1.pptx', '/svg4/mysql.svg');
            await pptToSvg('p2.pptx', '/svg3/mysql.svg');
            await flexibleDocxConvert('mysql.pptx',
                'mysql1.pdf');

            await flexibleDocxConvert('factsheet_Arabic.docx', 'factsheet_Arabic.pdf');
        } catch (err) {
            console.log(err);
        }

        console.log('Done.');
    }

    PDFNet.runWithCleanup(main).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
    }).then(function () { PDFNet.shutdown(); });
}



runOfficeToPDF();