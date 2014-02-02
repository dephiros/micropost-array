//image have to be named cell.tif, bottom.tif, and top.tif
//a cell in cell.tif have to be selected before running the macro
macro "Crop and Save [x]" {
    setBatchMode(true);
    postfix = " "
    if (isOpen("cell.tif")) {
	selectImage("cell.tif");
	postfix = ".tif";
    } else if (isOpen("cell.tiff")) {
	selectImage("cell.tiff");
	postfix = ".tiff"
    } else {
	exit("cell.tif is not opened");
    }
    getSelectionBounds(x, y, width, height);
    if (x = =0 && y == 0) {
	exit("cell is not selected");
    }

    //Create new directory for new image
    currentDir = getDirectory("image");
    temp = currentDir + "cell";
    for(i = 1; i > 0; i++) {
	myDirectory = temp + i + File.separator;
	if (!File.exists(myDirectory)) {
	    File.makeDirectory(myDirectory);
	    i = -1;
	}
    }
    name = "cell";
    run("Crop");
    saveAs("tiff", myDirectory + name + postfix);

    name = "top";
    selectImage(name + postfix);
    makeRectangle(x, y, width, height);
    run("Crop");
    saveAs("tiff", myDirectory + name + postfix);

    name = "bottom";
    selectImage(name + postfix);
    makeRectangle(x, y, width, height);
    run("Crop");
    saveAs("tiff", myDirectory + name + postfix);
    setBatchMode(false);
}
