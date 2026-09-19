/**********************************************************
 
Save as PNGs.jsx

DESCRIPTION
This sample gets files specified by the user from the 
selected folder and batch processes them and saves them 
as PNGs in the user desired destination with the same 
file name.

forked by Koichi Uchimura.
edited: 2020/10/12
 
**********************************************************/

// define variables.
var sourceFolder, files, sourceDoc, targetFile;

// Select the source folder.
sourceFolder = Folder.selectDialog('Select the folder with Illustrator files you want to convert to PNG', '~');

// Call main function.
main(sourceFolder, [".ai", ".eps"]);

/*********************************************************
main: main Function of PNGs.
**********************************************************/

function main(folderObj, ext) {
	// If a valid folder is selected
	if (folderObj != null) {
        var fileList = new Array();
    	fileList = folderObj.getFiles();
        if (fileList.length > 0) {
            getFolder(folderObj);
        }
        else {
            alert('No matching files found');
        }
	}
    function getFolder(folderObj) {
    	var fileList = folderObj.getFiles();
    	for (var i=0; i<fileList.length; i++){
    		if (fileList[i].getFiles) {
    			getFolder(fileList[i]); // サブフォルダがある限り繰り返す
    		} else {
    			var f = fileList[i].name.toLowerCase();
    			for(var j=0; j<ext.length; j++){
    				if (f.indexOf(ext[j]) > -1) { createPNG(fileList[i]); }
    			}
    		}
    	}
    }
}

/*********************************************************
createPNG: Function to chage target file to png format.
**********************************************************/

function createPNG(file) {
    app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
    sourceDoc = app.open(file); // returns the document object
    // Call function getNewName to get the name and file to save the PNG
    targetFile = getNewName();
    // Export as PNG
    sourceDoc.exportFile(targetFile, ExportType.PNG24, getPNGOptions());
    sourceDoc.close(SaveOptions.DONOTSAVECHANGES);
}

/*********************************************************
getNewName: Function to get the new file name. The primary
name is the same as the source file.
**********************************************************/

function getNewName() {
    var ext, docName, newName, saveInFile, docName;
    docName = sourceDoc.name;
    ext = '.png'; // new extension for file
    newName = "";

    for (var i = 0; docName[i] != "."; i++) {
        newName += docName[i];
    }
    newName += ext; // full name of the file

    // Create a file object to save the png
    saveInFile = new File(sourceFolder + '/' + newName);


    return saveInFile;
}

/*********************************************************
getPNGOptions: Function to set the PNG saving options of the
files using the ExportOptionsPNG24 object.
**********************************************************/

function getPNGOptions() {
    // Create the ExportOptionsPNG24 object to set the PNG options
    var opts = new ExportOptionsPNG24();

    // Setting ExportOptionsPNG24 properties. Please see the JavaScript Reference
    // for a description of these properties.
    // Add more properties here if you like
    opts.antiAliasing = true;
    opts.artBoardClipping = true;
    opts.horizontalScale = 100.0;
    opts.saveAsHTML = false;
    opts.transparency = true;
    opts.verticalScale = 100.0;

    return opts;
}