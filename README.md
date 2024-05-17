# DimensionCheckTable - AutoCAD Sample Code - Use at own risk
See also attached <a href="DimensionCheckTable.mp4">video</a> (you need to download it)

Given a layer named "DimensionsWithBalloons" with a table on it and some dimensions in the drawing area, you can call the command "DimensionCheckTable" and you will be prompted to select dimensions.
After pressing enter, the following things have been achieved:
- copies of the selected dimensions are created on the layer "DimensionsWithBalloons"
- position numbers are placed close to that dimensions on the layer "DimensionsWithBalloons"
- the table that was found on the layer "DimensionsWithBalloons" gets new rows appended on the bottom, each row for each dimension, where the dimension value and the max/min values based on tolerances will be inserted together with the position number
