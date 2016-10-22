# ColorBrewer Office Addin
A Microsoft Office COM add-in which allows the user to quickly change graph color schemes to a ColorBrewer palette.

![new picture 3](https://cloud.githubusercontent.com/assets/16891177/16711719/c9b65bf6-4627-11e6-9a93-ff371f6de55e.png)

## The v6.0 release of the ColorBrewer Office Add-in features:

### Compatibility across the basic suite of Microsoft Office applications
  * Excel 2007 and later
  * Word 2007 and later
  * Powerpoint 2007 and later

### Change the fill or line color of a chart by selecting a ColorBrewer palette from the drop-down gallery. Basic chart types supported in this release include:
  * **Scatter-type charts**
    * Scatterplots
    * Lines w/ markers
    * Radar w/ markers
  * **Line-type charts**
    * Lines w/o markers
    * Radar w/o markers
  * **Area-type charts** 
    * Area
    * Bar
    * Bubble
    * Column
    * Cone
    * Cylinder
    * Pyramid
    * Radar (filled)
  * **Pie-type charts**
    * Pie
    * Doughnut
    * Exploded Doughnut

  * *Notes:*
    * *The ColorBrewer Office Add-In supports variations on the major chart types, such as 3D, Stacked, and 100% Stacked.*
    * *Each ColorBrewer palette has a minimum (3) and maximum (usually 9, but some go as high as 12) number of colors. These ranges are specified next to the palette names on the drop-down menu. If the numbers of series in your chart is outside this range for the selected ColorBrewer palette, you won't be able to change the colors of that plot. If your chart has too many series for your desired color scheme, try a different palette with a greater maximum, and/or reduce your chart's series count.*

### Reverse the order of fill/line colors in a chart (see chart list above for supported types). Works with any number of series.
![new picture 4](https://cloud.githubusercontent.com/assets/16891177/16711769/0eeaceb2-462a-11e6-9255-5a5c788d61be.png)

----
#### Installation instructions and requirements:
***Note: Requires Microsoft Office 2007 or later, running on Windows.***

To install the ColorBrewer Office Add-in, simply run `ColorBrewerAddin.msi` and follow the prompts. To confirm that the add-in installed correctly, open up Excel, Word, or PowerPoint and look for the `ColorBrewer` ribbon tab.

To uninstall the add-in :cry:, navigate to the `Programs and Features` menu on your computer's Control Panel, look for `ColorBrewerAddin`, and click 'Uninstall'.
