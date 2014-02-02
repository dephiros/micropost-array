//image have to be named cell.tif, bottom.tif, and top.tif
//a cell in cell.tif have to be selected before running the macro
macro "Crop and Save [x]" {
    postfix = " "
    x = 0;
    y = 0;
    if (isOpen("cell.tif")) {
        postfix = ".tif";
    } else if (isOpen("cell.tiff")) {
        postfix = ".tiff"
    } else {
        exit("cell.tif is not opened");
    }
    currentDir = getDirectory("image");

    if (!isOpen("top" + postfix)) {
        if (File.exists(currentDir + "top" + postfix)) { 
            open(currentDir + "top" + postfix);
        } else {
            exit("top" + postfix + " does not exist");
        }
    }

    if (!isOpen("bottom" + postfix)) {
        if (File.exists(currentDir + "bottom" + postfix)) { 
            open(currentDir + "bottom" + postfix);
        } else {
            exit("bottom" + postfix + " does not exist");
        }
    }

    selectImage("cell" + postfix);
    getSelectionBounds(x, y, width, height);
    print(x);
    print(y);
    if (x ==0 && y == 0) {
        exit("cell is not selected");
    }

    //Create new directory for new image
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
}
