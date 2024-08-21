// ImageCatalogForLabel.jsx
// Fork of ImageCatalog.jsx
// Windows only

main();

// main
function main() {
	// Target files
	var myFilteredFiles;
	// Make certain that user interaction (display of dialogs, etc.) is turned on.
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
	// Display the folder browser. 
	var myFile = File.openDialog("Select the jpg file", ""); 
	// Get the path to the folder containing the files you want to place. 
	if(myFile != null){ 
		myFilteredFiles = myWinOSFileFilter(myFile);	
		if(myFilteredFiles.length != 0){ 
			myDisplayDialog(myFilteredFiles, myFile); 
			alert("Done!");
		} 
	}
}
// Windows version of the file filter.
function myWinOSFileFilter(myFile){
	// total labels
	var myFileCount = 12;
	// files
	var myFilteredFiles = new Array; 
	// loop for layout
	for(var myFileCounter = 0; myFileCounter < myFileCount; myFileCounter++){ 
		myFilteredFiles.push(myFile); 
	} 	
	return myFilteredFiles;
}

// Windows version of dialog
function myDisplayDialog(myFiles, myFile){ 
	var myLabelWidth = 112; 
	var myStyleNames = myGetParagraphStyleNames(app);
	var myLayerNames = ["Layer 1", "Labels"];
	var myDialog = app.dialogs.add({name:"Image Catalog"}); 
	with(myDialog.dialogColumns.add()){
		with(dialogRows.add()){ 
			staticTexts.add({staticLabel:"Information:"}); 
		}
		with(borderPanels.add()){
			with(dialogColumns.add()){
				with(dialogRows.add()){ 
					staticTexts.add({staticLabel:"Source Folder:", minWidth:myLabelWidth}); 
					staticTexts.add({staticLabel:myFile.path + "/" + myFile.name}); 
				} 
				with(dialogRows.add()){ 
					staticTexts.add({staticLabel:"Number of Images:", minWidth:myLabelWidth}); 
					staticTexts.add({staticLabel:myFiles.length + ""}); 
				} 
			}
		}
		with(dialogRows.add()){ 
				staticTexts.add({staticLabel:"Options:"}); 
		} 
		with(borderPanels.add()){
			with(dialogColumns.add()){
				with(dialogRows.add()){
					staticTexts.add({staticLabel:"Number of Rows:", minWidth:myLabelWidth}); 
					var myNumberOfRowsField = integerEditboxes.add({editValue:3}); 
				} 
				with(dialogRows.add()){ 
					staticTexts.add({staticLabel:"Number of Columns:", minWidth:myLabelWidth}); 
					var myNumberOfColumnsField = integerEditboxes.add({editValue:3}); 
				} 
				with(dialogRows.add()){ 
					staticTexts.add({staticLabel:"Horizontal Offset:", minWidth:myLabelWidth}); 
					var myHorizontalOffsetField = measurementEditboxes.add({editValue:12, editUnits:MeasurementUnits.points}); 
				} 
				with(dialogRows.add()){ 
					staticTexts.add({staticLabel:"Vertical Offset:", minWidth:myLabelWidth}); 
					var myVerticalOffsetField = measurementEditboxes.add({editValue:24, editUnits:MeasurementUnits.points}); 
				} 
				with (dialogRows.add()){
					with(dialogColumns.add()){
						staticTexts.add({staticLabel:"Fitting:", minWidth:myLabelWidth}); 
					}
					with(dialogColumns.add()){
						var myFitProportionalCheckbox = checkboxControls.add({staticLabel:"Proportional", checkedState:true});
						var myFitCenterContentCheckbox = checkboxControls.add({staticLabel:"Center Content", checkedState:true});
						var myFitFrameToContentCheckbox = checkboxControls.add({staticLabel:"Frame to Content", checkedState:true});
					}
				}
				with(dialogRows.add()){ 
						var myRemoveEmptyFramesCheckbox = checkboxControls.add({staticLabel:"Remove Empty Frames:", checkedState:true}); 
				}
			}
		}
		with(dialogRows.add()){ 
				staticTexts.add({staticLabel:""}); 
		} 
		var myLabelsGroup = enablingGroups.add({staticLabel:"Labels", checkedState:true});
		with (myLabelsGroup){
			with(dialogColumns.add()){
				//Label type
				with(dialogRows.add()){
					with(dialogColumns.add()){
						staticTexts.add({staticLabel:"Label Type:", minWidth:myLabelWidth});
					}
					with(dialogColumns.add()){
						var myLabelTypeDropdown = dropdowns.add({stringList:["File name", "File path", "XMP description", "XMP author"], selectedIndex:0});
					}
				}
				//Text frame height
				with(dialogRows.add()){
					with(dialogColumns.add()){
						staticTexts.add({staticLabel:"Label Height:", minWidth:myLabelWidth});
					}
					with(dialogColumns.add()){
						var myLabelHeightField = measurementEditboxes.add({editValue:24, editUnits:MeasurementUnits.points});
					}
				}
				//Text frame offset
				with(dialogRows.add()){
					with(dialogColumns.add()){
						staticTexts.add({staticLabel:"Label Offset:", minWidth:myLabelWidth});
					}
					with(dialogColumns.add()){
						var myLabelOffsetField = measurementEditboxes.add({editValue:0, editUnits:MeasurementUnits.points});
					}
				}
				//Style to apply
				with(dialogRows.add()){
					with(dialogColumns.add()){
						staticTexts.add({staticLabel:"Label Style:", minWidth:myLabelWidth});
					}
					with(dialogColumns.add()){
						var myLabelStyleDropdown = dropdowns.add({stringList:myStyleNames, selectedIndex:0});
					}
				}
				//Layer
				with(dialogRows.add()){
					with(dialogColumns.add()){
						staticTexts.add({staticLabel:"Layer:", minWidth:myLabelWidth});
					}
					with(dialogColumns.add()){
						var myLayerDropdown = dropdowns.add({stringList:myLayerNames, selectedIndex:0});
					}
				}
			}
		}
        var myResult = myDialog.show(); 
        if(myResult == true){ 
			var myNumberOfRows = myNumberOfRowsField.editValue; 
			var myNumberOfColumns = myNumberOfColumnsField.editValue; 
			var myRemoveEmptyFrames = myRemoveEmptyFramesCheckbox.checkedState;
			var myFitProportional = myFitProportionalCheckbox.checkedState;
			var myFitCenterContent = myFitCenterContentCheckbox.checkedState;
			var myFitFrameToContent = myFitFrameToContentCheckbox.checkedState;
			var myHorizontalOffset = myHorizontalOffsetField.editValue;
			var myVerticalOffset = myVerticalOffsetField.editValue;
			var myMakeLabels = myLabelsGroup.checkedState;
			var myLabelType = myLabelTypeDropdown.selectedIndex;
			var myLabelHeight = myLabelHeightField.editValue;
			var myLabelOffset = myLabelOffsetField.editValue;
			var myLabelStyle = myStyleNames[myLabelStyleDropdown.selectedIndex];
			var myLayerName = myLayerNames[myLayerDropdown.selectedIndex];
			myDialog.destroy();
			myMakeImageCatalog(myFiles, myNumberOfRows, myNumberOfColumns, myRemoveEmptyFrames, myFitProportional, myFitCenterContent, myFitFrameToContent, myHorizontalOffset, myVerticalOffset, myMakeLabels, myLabelType, myLabelHeight, myLabelOffset, myLabelStyle,  myLayerName); 
        } 
		else{
			myDialog.destroy(); 
		}
	}
}

function myGetParagraphStyleNames(myDocument){
	var myStyleNames = new Array;
	var myAddLabelStyle = true;
	for(var myCounter = 0; myCounter < myDocument.paragraphStyles.length; myCounter++){
		myStyleNames.push(myDocument.paragraphStyles.item(myCounter).name);
		if (myDocument.paragraphStyles.item(myCounter).name == "Labels"){
			myAddLabelStyle = false;
		}
	}
	if(myAddLabelStyle == true){
		myStyleNames.push("Labels");
	}
	return myStyleNames;
}

function myMakeImageCatalog(myFiles, myNumberOfRows, myNumberOfColumns, myRemoveEmptyFrames, myFitProportional, myFitCenterContent, myFitFrameToContent, myHorizontalOffset, myVerticalOffset, myMakeLabels, myLabelType, myLabelHeight, myLabelOffset, myLabelStyle,  myLayerName){
	var myPage, myFile, myCounter, myX1, myY1, myX2, myY2, myRectangle, myLabelStyle, myLabelLayer; 
	var myParagraphStyle, myError;
	var myFramesPerPage = myNumberOfRows * myNumberOfColumns;  
	var myDocument = app.open("C:\\Users\\ebisu-do\\Documents\\work\\auto\\test11.indd");
	myDocument.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.points; 
	myDocument.viewPreferences.verticalMeasurementUnits = MeasurementUnits.points; 
	var myDocumentPreferences = myDocument.documentPreferences;
	var myNumberOfFrames = myFiles.length; 
	var myNumberOfPages = Math.round(myNumberOfFrames / myFramesPerPage); 
	if ((myNumberOfPages * myFramesPerPage) < myNumberOfFrames){  
		myNumberOfPages++; 
	}
	//If myMakeLabels is true, then add the label style and layer if they do not already exist.
	if(myMakeLabels == true){
		try{
			myLabelLayer = myDocument.layers.item(myLayerName);
			//if the layer does not exist, trying to get the layer name will cause an error.
			myLabelLayer.name;
		}
		catch (myError){
			myLabelLayer = myDocument.layers.add({name:myLayerName}); 
		}
		//If the paragraph style does not exist, create it.
		try{
			myParagraphStyle = myDocument.paragraphStyles.item(myLabelStyle);
			myParagraphStyle.name;
		}
		catch(myError){
			myDocument.paragraphStyles.add({name:myLabelStyle});
		}
	}
	myDocumentPreferences.pagesPerDocument = myNumberOfPages;  
	myDocumentPreferences.facingPages = false;  
	var myPage = myDocument.pages.item(0);  
	var myMarginPreferences = myPage.marginPreferences; 
	var myLeftMargin = myMarginPreferences.left;  
	var myTopMargin = myMarginPreferences.top;  
	var myRightMargin = myMarginPreferences.right;  
	var myBottomMargin = myMarginPreferences.bottom;  
	var myLiveWidth = (myDocumentPreferences.pageWidth - (myLeftMargin + myRightMargin)) + myHorizontalOffset
	var myLiveHeight = myDocumentPreferences.pageHeight - (myTopMargin + myBottomMargin)
	var myColumnWidth = myLiveWidth / myNumberOfColumns
	var myFrameWidth = myColumnWidth - myHorizontalOffset
	var myRowHeight = (myLiveHeight / myNumberOfRows)
	var myFrameHeight = myRowHeight - myVerticalOffset
	var myPages = myDocument.pages;
	// Construct the frames in reverse order. Don't laugh--this will  
	// save us time later (when we place the graphics).  
	for (myCounter = myDocument.pages.length-1; myCounter >= 0; myCounter--){  
		myPage = myPages.item(myCounter); 
		for (var myRowCounter = myNumberOfRows; myRowCounter >= 1; myRowCounter--){  
			myY1 = myTopMargin + (myRowHeight * (myRowCounter-1));
			myY2 = myY1 + myFrameHeight;
			for (var myColumnCounter = myNumberOfColumns; myColumnCounter >= 1; myColumnCounter--){  
				myX1 = myLeftMargin + (myColumnWidth * (myColumnCounter-1));
				myX2 = myX1 + myFrameWidth;
				myRectangle = myPage.rectangles.add(myDocument.layers.item(-1), undefined, undefined, {geometricBounds:[myY1, myX1, myY2, myX2], strokeWeight:0, strokeColor:myDocument.swatches.item("None")});  
			} 
		}
	}
	// Because we constructed the frames in reverse order, rectangle 1  
	// is the first rectangle on page 1, so we can simply iterate through  
	// the rectangles, placing a file in each one in turn. myFiles = myFolder.Files;  
	for (myCounter = 0; myCounter < myNumberOfFrames; myCounter++){  
		myFile = myFiles[myCounter];  
		myRectangle = myDocument.rectangles.item(myCounter); 
		myRectangle.place(File(myFile)); 
		myRectangle.label = myFile.fsName.toString();
		//Apply fitting options as specified.
		if(myFitProportional){
			myRectangle.fit(FitOptions.proportionally); 
		}
		if(myFitCenterContent){
			myRectangle.fit(FitOptions.centerContent); 
		}
		if(myFitFrameToContent){
			myRectangle.fit(FitOptions.frameToContent); 
		}
		//Add the label, if necessary.
		if(myMakeLabels == true){ 
			myAddLabel(myRectangle, myLabelType, myLabelHeight, myLabelOffset, myLabelStyle, myLayerName);
		} 
	}
	if (myRemoveEmptyFrames == 1){  
		for (var myCounter = myDocument.rectangles.length-1; myCounter >= 0;myCounter--){  
			if (myDocument.rectangles.item(myCounter).contentType == ContentType.unassigned){ 
				myDocument.rectangles.item(myCounter).remove(); 
			}  
			else{ 
				//As soon as you encounter a rectangle with content, exit the loop. 
				break; 
			} 
		}
	}
}

// Function that adds the label.
function myAddLabel(myFrame, myLabelType, myLabelHeight, myLabelOffset, myLabelStyleName, myLayerName){
	var myDocument = app.documents.item(0);
	var myLabel;
	var myLabelStyle = myDocument.paragraphStyles.item(myLabelStyleName);
	var myLabelLayer = myDocument.layers.item(myLayerName);
	var myLink =myFrame.graphics.item(0).itemLink;
	//Label type defines the text that goes in the label.
	switch(myLabelType){
		//File name
		case 0:
			myLabel = myLink.name;
			break;
		//File path
		case 1:
			myLabel = myLink.filePath;
			break;
		//XMP description
		case 2:
			try{
				myLabel = myLink.linkXmp.description;
				if(myLabel.replace(/^\s*$/gi, "")==""){
					throw myError;
				}
			}
			catch(myError){
				myLabel = "No description available.";
			}
			break;
		//XMP author
		case 3:
			try{
				myLabel = myLink.linkXmp.author
				if(myLabel.replace(/^\s*$/gi, "")==""){
					throw myError;
				}
			}
			catch(myError){
				myLabel = "No author available.";
			}
			break;
	}
	var myX1 = myFrame.geometricBounds[1]; 
	var myY1 = myFrame.geometricBounds[2] + myLabelOffset; 
	var myX2 = myFrame.geometricBounds[3]; 
	var myY2 = myY1 + myLabelHeight;
	var myTextFrame = myFrame.parent.textFrames.add(myLabelLayer, undefined, undefined,{geometricBounds:[myY1, myX1, myY2, myX2], contents:myLabel}); 
	myTextFrame.textFramePreferences.firstBaselineOffset = FirstBaseline.leadingOffset; 
	myTextFrame.parentStory.texts.item(0).appliedParagraphStyle = myLabelStyle;				
}
