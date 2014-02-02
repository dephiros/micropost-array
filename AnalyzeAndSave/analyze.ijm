// check to see if image is binary.
// if yes then set appropriate value and run analyze particle on the image
macro "Analyze and Save[a]" {
      // check if the current image is binary
      getHistogram(values, counts, 256); 
      for (i=1;i<255;i++) total+=counts[i]; 
      if (total>0) exit("8-bit binary image (0 and 255) required.");

      // get the path and name of current image
      name = getInfo("image.filename");
      path = getInfo("image.directory");
      savedName = "";
      if (matches(name, ".*[b|B]ottom.*")) savedName = "bottom.xls";
      if (matches(name, ".*[t|T]op.*")) savedName = "top.xls";

      // Set the appropriate measurement
      run("Set Measurements...", "area mean min centroid center fit invert redirect=None decimal=3");
      run("Analyze Particles...", "size=0-Infinity circularity=0.00-1.00 show=Outlines display clear");

      // Save the result
      saveAs("Results", path + savedName);

}
