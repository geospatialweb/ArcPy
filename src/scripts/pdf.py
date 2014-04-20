# ----------------
# Import required modules
# ----------------
import arcpy, os, win32com.client as com

# ----------------
# Set output scratch directory and geodatabase workspaces
# ----------------
scratch_dir = arcpy.env.scratchWorkspace
scratch_gdb = os.path.join(scratch_dir, "scratch.gdb")

# ----------------
# Set input data folder paths
# ----------------
inFolder = r"C:\data"
inFolderToolData = os.path.join(inFolder, "ToolData")
inFolderSpatialRef = os.path.join(inFolder, "SpatialReference")
inFolderOntarioGDB = os.path.join(inFolder, "Ontario.gdb")

# ----------------
# Set references to derived layer symbology
# ----------------
inPointLayer = os.path.join(inFolderToolData, "SitePoint.lyr")
inBuffer1Layer = os.path.join(inFolderToolData, "Buffer1.lyr")
inBuffer2Layer = os.path.join(inFolderToolData, "Buffer2.lyr")

# ----------------
# Set reference to transparent Ontario UTM 10km Grid to later determine correct NAD83 UTM Zone spatial reference
# ----------------
inUTM10kmGridLayer = os.path.join(inFolderToolData, "UTM10kmGrid.lyr")

# ----------------
# Set scratch workspace feature class references
# ----------------
outPointFCPath = os.path.join(scratch_gdb, "SitePoint")
outBuffer1FCPath = os.path.join(scratch_gdb, "Buffer1")
outBuffer2FCPath = os.path.join(scratch_gdb, "Buffer2")

# ----------------
# Set Geoprocessing input parameters to local variables
# Note: Index values must match GP Service
# ----------------
mapType = arcpy.GetParameterAsText(0)
orderID = arcpy.GetParameterAsText(1)
xCoord = float(arcpy.GetParameterAsText(2))
yCoord = float(arcpy.GetParameterAsText(3))
mapScale = int(arcpy.GetParameterAsText(4))
buffer1Distance = int(arcpy.GetParameterAsText(5))
buffer2Distance = int(arcpy.GetParameterAsText(6))

# ----------------
# Manipulate the Map Surround
# ----------------
# Get map document. Based on the choice of map, choose the correct mxd...
# Choices are "OBM", "Soils", "Physiography", "ANSI", "Surface Geology" or "Bedrock Geology"
if mapType == "OBM":
    mxd = arcpy.mapping.MapDocument(os.path.join(inFolder, "OBMLayout.mxd"))
elif mapType == "Soils":
    mxd = arcpy.mapping.MapDocument(os.path.join(inFolder, "SoilsLayout.mxd"))
elif mapType == "Physiography":
    mxd = arcpy.mapping.MapDocument(os.path.join(inFolder, "PhysiographyLayout.mxd"))
elif mapType == "ANSI":
    mxd = arcpy.mapping.MapDocument(os.path.join(inFolder, "ANSILayout.mxd"))
elif mapType == "Surface Geology":
    mxd = arcpy.mapping.MapDocument(os.path.join(inFolder, "SurfaceGeologyLayout.mxd"))
elif mapType == "Bedrock Geology":
    mxd = arcpy.mapping.MapDocument(os.path.join(inFolder, "BedrockGeologyLayout.mxd"))

# Set reference to mxd data frame
df = arcpy.mapping.ListDataFrames(mxd)[0]

# ----------------
# Set Site Point
# ----------------
# Create point feature class for site point
arcpy.CreateFeatureclass_management(scratch_gdb, os.path.basename(outPointFCPath), "Point", "", "DISABLED", "DISABLED", arcpy.SpatialReference(os.path.join(inFolderSpatialRef, "GCS North American 1983.prj")))

# Create point geometry object (shape) from point coordinate object
point = arcpy.PointGeometry(arcpy.Point(xCoord, yCoord))

# Create point feature class insert cursor
cur = arcpy.InsertCursor(outPointFCPath)

# Advance to first record
row = cur.newRow()

# Assign point geometry to feature class record 
row.shape = point

# Insert feature class record
cur.insertRow(row)

# Delete cursor and row objects to release locks on feature class
del cur, row

# Add fields POINT_X & POINT_Y (xCoord, yCoord) as record attributes
arcpy.AddXY_management(outPointFCPath)

# Create point layer object
outPointLayer = arcpy.mapping.Layer(outPointFCPath)

# Add point layer object to top of mxd data frame TOC
arcpy.mapping.AddLayer(df, outPointLayer, "TOP")

# Assign layer symbology
arcpy.ApplySymbologyFromLayer_management(outPointLayer, inPointLayer)

# ----------------
# Set Buffer1
# ----------------
# Create buffer feature class
arcpy.Buffer_analysis(outPointFCPath, outBuffer1FCPath, str(buffer1Distance) + " Meters")

# Create buffer layer object
outBuffer1Layer = arcpy.mapping.Layer(outBuffer1FCPath)

# Insert buffer layer object into mxd data frame
arcpy.mapping.InsertLayer(df, outPointLayer, outBuffer1Layer, "AFTER")

# ----------------
# Set Buffer2
# ----------------
# Create buffer feature class
arcpy.Buffer_analysis(outPointFCPath, outBuffer2FCPath, str(buffer2Distance) + " Meters")

# Create buffer layer object
outBuffer2Layer = arcpy.mapping.Layer(outBuffer2FCPath)

# Insert buffer layer object into mxd data frame
arcpy.mapping.InsertLayer(df, outBuffer1Layer, outBuffer2Layer, "AFTER")

# ----------------
# Set mxd data frame spatial reference
# ----------------
# Create UTM layer object
utm10kmGridLayer = arcpy.mapping.Layer(inUTM10kmGridLayer)

# Select polygon in UTM layer that contains site point
arcpy.SelectLayerByLocation_management(utm10kmGridLayer, "CONTAINS", outPointLayer)

# Create UTM layer search cursor
cur = arcpy.SearchCursor(utm10kmGridLayer)

# Advance to first record
row = cur.next()

# Set variable to UTM_ZONE field value in selected record
utmZone = row.getValue("UTM_ZONE")

# Delete cursor and row objects to release locks on UTM Layer
del cur, row

# Set mxd data frame spatial reference to selected NAD83 UTM Zone
df.spatialReference = arcpy.SpatialReference(os.path.join(inFolderSpatialRef, "NAD 1983 UTM Zone " + str(utmZone) + "N.prj"))

# ----------------
# Set Map Surround layout properties
# ----------------
# Set map scale to input parameter
df.scale = mapScale

# Depending on which buffer layer has smaller buffer distance...
if buffer1Distance <= buffer2Distance:
    # Assign buffer layer symbology
    arcpy.ApplySymbologyFromLayer_management(outBuffer1Layer, inBuffer1Layer)
    arcpy.ApplySymbologyFromLayer_management(outBuffer2Layer, inBuffer2Layer)
    
    # Pan to buffer layer extent
    df.panToExtent(outBuffer2Layer.getExtent())
else:
    # Assign buffer layer symbology
    arcpy.ApplySymbologyFromLayer_management(outBuffer1Layer, inBuffer2Layer)
    arcpy.ApplySymbologyFromLayer_management(outBuffer2Layer, inBuffer1Layer)
    
    # Pan to buffer layer extent
    df.panToExtent(outBuffer1Layer.getExtent())

# Get reference to buffer1 distance text element
buffer1TextElement = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "buffer1DistanceText")[0]

# Change buffer1 distance text in mxd to reflect current Order ID Number
buffer1TextElement.text = str(buffer1Distance) + "m"

# Get reference to buffer2 distance text element
buffer2TextElement = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "buffer2DistanceText")[0]

# Change buffer2 distance text in mxd to reflect current Order ID Number
buffer2TextElement.text = str(buffer2Distance) + "m"

# Get reference to spatial reference text element
projCoordTextElement = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "projCoordText")[0]

# Change spatial reference text in mxd to reflect current spatial reference
projCoordTextElement.text = "NAD 1983 UTM Zone " + str(utmZone) + "N"

# Get reference to Order ID text element
orderIDTextElement = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "orderIDText")[0]

# Change Order ID text in mxd to reflect current Order ID number
orderIDTextElement.text = "Order No. " + orderID

# ----------------
# Generate Report PDF
# ----------------
# Only SOILS map type requires Report PDF
if mapType == "Soils":
    # Set output clip feature class path
    outBufferClipFCPath = os.path.join(scratch_gdb, "BufferClip")
    
    # Set buffer layer to 'largest' buffer distance
    if buffer1Distance <= buffer2Distance:
        outBufferLayer = outBuffer2Layer
    else:
        outBufferLayer = outBuffer1Layer
    
    # Create 2000m buffer clip feature class
    arcpy.Clip_analysis(os.path.join(inFolderOntarioGDB, "Soils"), outBufferLayer, outBufferClipFCPath)
    
    # Create clip feature class search cursor
    cur = arcpy.SearchCursor(outBufferClipFCPath)
    
    # Advance to first record and initialize row counter
    row = cur.next()
    rowNum = 1
    
    # Create Excel application, workbook and worksheet objects
    xlApp = com.Dispatch("Excel.Application")
    xlApp.Visible = 0
    xlBook = xlApp.Workbooks.Add()
    xlSheet = xlBook.ActiveSheet
    
    # Set page orientation to landscape
    xlSheet.PageSetup.Orientation = com.constants.xlLandscape
    
    # Set worksheet column widths
    xlSheet.Columns(1).ColumnWidth = 30
    xlSheet.Columns(2).ColumnWidth = 13
    xlSheet.Columns(3).ColumnWidth = 13
    xlSheet.Columns(4).ColumnWidth = 35
    xlSheet.Columns(5).ColumnWidth = 8
    
    # Process all rows in clip feature class attribute table
    while row:
        # Output group report headers
        xlSheet.Range(xlSheet.Cells(rowNum,1), xlSheet.Cells(rowNum,3)).Font.Bold = True
        xlSheet.Range(xlSheet.Cells(rowNum,1), xlSheet.Cells(rowNum,3)).Font.Underline = True
        xlSheet.Cells(rowNum,1).Value = "MAPUNIT"
        xlSheet.Cells(rowNum,2).Value = "SOIL COMPLEX"
        xlSheet.Cells(rowNum,3).Value = "AREA (hectares)"
        
        # Output attribute table values
        rowNum += 1
        xlSheet.Cells(rowNum,2).HorizontalAlignment = -4108
        xlSheet.Cells(rowNum,1).Value = row.MAPUNIT
        xlSheet.Cells(rowNum,2).Value = row.SOIL_CMPLX
        xlSheet.Cells(rowNum,3).Value = "%.2f" % (row.Area_m / 10000)
        
        # Output secondary report headers
        rowNum += 1
        xlSheet.Range(xlSheet.Cells(rowNum,1), xlSheet.Cells(rowNum,5)).Font.Bold = True
        xlSheet.Range(xlSheet.Cells(rowNum,1), xlSheet.Cells(rowNum,5)).Font.Underline = True
        xlSheet.Cells(rowNum,1).HorizontalAlignment = -4152
        xlSheet.Cells(rowNum,1).Value = "PERCENT"
        xlSheet.Cells(rowNum,2).Value = "SOIL TYPE"
        xlSheet.Cells(rowNum,3).Value = "SOIL CODE"
        xlSheet.Cells(rowNum,4).Value = "SOIL NAME"
        xlSheet.Cells(rowNum,5).Value = "SYMBOL"
        
        # Output attribute table values
        rowNum += 1
        xlSheet.Cells(rowNum,1).HorizontalAlignment = -4152
        xlSheet.Cells(rowNum,1).Value = row.PERCENT1
        xlSheet.Cells(rowNum,2).Value = row.SOILTYPE1
        xlSheet.Cells(rowNum,3).Value = row.SOILCODE1
        xlSheet.Cells(rowNum,4).Value = row.SOIL_NAME1
        xlSheet.Cells(rowNum,5).Value = row.SYMBOL1
        
        # SOIL_CMPLX value determines number of soil data report rows (1 min to 3 max) per clip feature class record
        if row.SOIL_CMPLX != 1:
            # Output attribute table values
            rowNum += 1
            xlSheet.Cells(rowNum,1).HorizontalAlignment = -4152
            xlSheet.Cells(rowNum,1).Value = row.PERCENT2
            xlSheet.Cells(rowNum,2).Value = row.SOILTYPE2
            xlSheet.Cells(rowNum,3).Value = row.SOILCODE2
            xlSheet.Cells(rowNum,4).Value = row.SOIL_NAME2
            xlSheet.Cells(rowNum,5).Value = row.SYMBOL2
            
            if row.SOIL_CMPLX == 3:
                # Output attribute table values
                rowNum += 1
                xlSheet.Cells(rowNum,1).HorizontalAlignment = -4152
                xlSheet.Cells(rowNum,1).Value = row.PERCENT3
                xlSheet.Cells(rowNum,2).Value = row.SOILTYPE3
                xlSheet.Cells(rowNum,3).Value = row.SOILCODE3
                xlSheet.Cells(rowNum,4).Value = row.SOIL_NAME3
                xlSheet.Cells(rowNum,5).Value = row.SYMBOL3
        
        # Advance to next record and increment row counter
        row = cur.next()
        rowNum += 2
    
    # Delete cursor and row objects to release locks on clip feature class
    del cur, row
    
    # Create PDF Creator application object and start application
    PDFCreator = com.Dispatch("PDFCreator.clsPDFCreator")
    PDFCreator.cStart("/NoProcessingAtStartup", 1)
    
    # Set specific PDF Creator printer driver properties
    options = PDFCreator.cOptions
    options.UseAutosave = 1
    options.UseAutosaveDirectory = 1
    options.AutosaveDirectory = scratch_dir
    options.AutosaveFilename = mapType
    options.AutosaveFormat = 0  # 0 = PDF
    PDFCreator.cOptions = options
    PDFCreator.cSaveOptions()
    PDFCreator.cSaveOptionsToFile(os.path.join(os.path.join(inFolder, "Logs"), "PDFCreator.ini"))
    
    # Clear printer cache and quit PDF Creator application
    PDFCreator.cClearCache()
    PDFCreator.cClose()
    
    # Generate Report PDF using PDF Creator printer driver
    xlSheet.PrintOut(1, 99, 1, False, "PDFCreator", False)
    
    # Close Excel workbook
    xlBook.Close(SaveChanges = 0)
    
    # Quit Excel application
    xlApp.Quit()

# ----------------
# Export PDF file
# ----------------
# Set Map PDF and Report PDF paths
outMapPDFPath = os.path.join(scratch_dir, mapType + " Report_Order " + orderID + ".pdf")
outReportPDFPath = os.path.join(scratch_dir, mapType + ".pdf")

# Generate Map PDF
arcpy.mapping.ExportToPDF(mxd, outMapPDFPath)

# Only SOILS map type requires Report PDF
if mapType == "Soils":
    # Open Map PDF
    pdfDoc = arcpy.mapping.PDFDocumentOpen(outMapPDFPath)
    
    # Append Report PDF
    pdfDoc.appendPages(outReportPDFPath)
    
    # Save Map PDF
    pdfDoc.saveAndClose()

# Pass Map PDF back to client as Geoprocessing output parameter with URL to Map PDF path
arcpy.SetParameterAsText(7, outMapPDFPath)

# Create KML layers
arcpy.LayerToKML_conversion(outPointLayer, os.path.join(scratch_dir, os.path.basename(inPointLayer)[:-4] + ".kmz"), mapScale, "COMPOSITE")
arcpy.LayerToKML_conversion(outBuffer1Layer, os.path.join(scratch_dir, os.path.basename(inBuffer1Layer)[:-4] + ".kmz"), mapScale, "COMPOSITE")
arcpy.LayerToKML_conversion(outBuffer2Layer, os.path.join(scratch_dir, os.path.basename(inBuffer2Layer)[:-4] + ".kmz"), mapScale, "COMPOSITE")

# ----------------
# Clean up... 
# ----------------
# Delete data variables
del scratch_dir, scratch_gdb, inFolder, inFolderToolData, inFolderSpatialRef, inFolderOntarioGDB, inPointLayer, inBuffer1Layer, inBuffer2Layer, inUTM10kmGridLayer, outPointFCPath, outBuffer1FCPath, outBuffer2FCPath, orderID, xCoord, yCoord, mapScale, buffer1Distance, buffer2Distance, mxd, df, point, outPointLayer, outBuffer1Layer, outBuffer2Layer, utm10kmGridLayer, utmZone, buffer1TextElement, buffer2TextElement, projCoordTextElement, orderIDTextElement, outMapPDFPath, outReportPDFPath
# Only SOILS map type requires Report PDF
if mapType == "Soils":
    del outBufferClipFCPath, outBufferLayer, rowNum, xlApp, xlBook, xlSheet, PDFCreator, options, pdfDoc, mapType