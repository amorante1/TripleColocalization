# TripleColocalization
Program used to calculate the colocalization between three color channels during immunofluoresence micoscopy. Program will calculate the percentage of overlap of channel 1 to both channels 2 and 3 together. Output colocalization is based on Manders' Colocalization Coefficient.
To use this file:

1. External libraries needed: PIL, xlwt, xlrd, xlutils
2. Create an empty folder anywhere on computer.
3. Input your three individual color channel images into the folder.
4. Start the file and, when prompted, select the folder that contains the three images to analyze.
5. Adjust the image brightness as necessary with slider bar and press Update button.
6. Press Analyze button.
7. Excel sheet should autogenerate in Colocalization folder with mM1, mM2, and mM3 coefficients. 
