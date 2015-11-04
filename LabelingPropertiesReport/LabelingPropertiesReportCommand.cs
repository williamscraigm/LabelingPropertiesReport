using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Windows.Forms;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.ArcMapUI;
using ESRI.ArcGIS.ArcMap;

namespace LabelingPropertiesReport
{
    public class LabelingPropertiesReportCommand : ESRI.ArcGIS.Desktop.AddIns.Button
    {
        public LabelingPropertiesReportCommand()
        {
        }

        protected override void OnClick()
        {
            IMxDocument mxd = ArcMap.Document as IMxDocument;
            IMaps maps = mxd.Maps;
            int mapCount = maps.Count;

            string docFolderPath = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string MapDocName = Path.GetFileNameWithoutExtension(ArcMap.Application.Document.Title);
            string folderPath = docFolderPath + "\\LabelingReport\\";
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            string path = folderPath + MapDocName + ".csv";
            if (File.Exists(path))
            {
                File.Delete(path);
            }

            StreamWriter sw = new StreamWriter(path, false);
            TextWriter tw = sw as TextWriter;
           
            sw.AutoFlush = true;

            tw.WriteLine(GetFieldColumns());

            for (int i = 0; i < mapCount; i++)
            {
                IMap map = maps.get_Item(i);
                ProcessMap(map, tw);
            }

            tw.Close();

            MessageBox.Show("Report generated at: " + path, "Labeling Report", MessageBoxButtons.OK);
            ArcMap.Application.CurrentTool = null;
        }
        private static void ProcessMap(IMap map, TextWriter tw)
        {
            IMapOverposter mapOverposter = map as IMapOverposter;
            IMaplexOverposterProperties maplexOverposterProps = mapOverposter.OverposterProperties as IMaplexOverposterProperties;

            //only handle Maplex maps
            if (maplexOverposterProps == null)
            {
                tw.WriteLine("Skipped Standard Label Engine Map: " + map.Name + ",");
                return;
            }

            ESRI.ArcGIS.esriSystem.UID flUID = new ESRI.ArcGIS.esriSystem.UIDClass();
            flUID.Value = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}"; //IGeoFeatureLayer
            IEnumLayer enumLayer = map.get_Layers(flUID, true);
            enumLayer.Reset();
            IGeoFeatureLayer featureLayer = enumLayer.Next() as IGeoFeatureLayer;

            while (featureLayer != null)
            {
                if (featureLayer.DisplayAnnotation == true)
                {
                    ProcessLayer(map.Name, featureLayer, tw);
                }

                featureLayer = enumLayer.Next() as IGeoFeatureLayer;
            }
        }
        private static void ProcessLayer(string mapName, IGeoFeatureLayer featureLayer, TextWriter tw)
        {
            IAnnotateLayerPropertiesCollection2 annoPC = featureLayer.AnnotationProperties as IAnnotateLayerPropertiesCollection2;
            int labelClassCount = annoPC.Count;
            IAnnotateLayerProperties annoLP = null;
            int classID = 0;

            for (int i = 0; i < labelClassCount; i++)
            {
                annoPC.QueryItem(i, out annoLP, out classID);
                if (annoLP.DisplayAnnotation == true)
                {
                    string labelClassString = ProcessLabelClass(mapName, featureLayer.Name, annoLP, classID);
                    tw.WriteLine(labelClassString);
                }
            }
        }
        private static string ProcessLabelClass(string mapName, string layerName, IAnnotateLayerProperties annoLP, int ID)
        {

            StringBuilder stringBuild = new StringBuilder();

            ILabelEngineLayerProperties2 labelEngineLP = annoLP as ILabelEngineLayerProperties2;

            stringBuild.Append(CSV.Escape(mapName)); //Map name
            stringBuild.Append("," + CSV.Escape(layerName)); //Layer name
            stringBuild.Append("," + CSV.Escape(annoLP.Class)); //Label class name
            // stringBuild.Append("," + CSV.Escape(labelEngineLP.Offset)); //Offset   //Not used at this point
            stringBuild.Append("," + CSV.Escape(annoLP.WhereClause)); //SQL query
            stringBuild.Append("," + CSV.Escape(annoLP.AnnotationMaximumScale)); //Visiblity max scale
            stringBuild.Append("," + CSV.Escape(annoLP.AnnotationMinimumScale)); //Visiblity min scale
            stringBuild.Append("," + CSV.Escape(labelEngineLP.ExpressionParser.Name)); //Label expression language
            stringBuild.Append("," + CSV.Escape(labelEngineLP.Expression)); //Label expression

            ICodedValueAttributes annoExpEngine = labelEngineLP.ExpressionParser as ICodedValueAttributes;
            stringBuild.Append("," + CSV.Escape(annoExpEngine.UseCodedValue)); //Use coded value description

            IMaplexOverposterLayerProperties maplexOverposter = labelEngineLP.OverposterLayerProperties as IMaplexOverposterLayerProperties;
            if (maplexOverposter != null)
            {
                #region Maplex Properties
                IMaplexOverposterLayerProperties2 maplexOverposter2 = maplexOverposter as IMaplexOverposterLayerProperties2;
                IMaplexOverposterLayerProperties3 maplexOverposter3 = maplexOverposter as IMaplexOverposterLayerProperties3;
                IMaplexOverposterLayerProperties4 maplexOverposter4 = maplexOverposter as IMaplexOverposterLayerProperties4;

                stringBuild.Append("," + CSV.Escape(maplexOverposter4.RemoveExtraLineBreaks)); //RemoveExtraLineBreaks
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.RemoveExtraWhiteSpace)); //RemoveExtraWhiteSpace

                //Placement points
                stringBuild.Append("," + CSV.Escape(maplexOverposter.PointPlacementMethod)); //Point placement method
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.UseExactSymbolOutline)); //Measure offset from exact symbol outline
                stringBuild.Append("," + CSV.Escape(maplexOverposter.CanShiftPointLabel)); //May shift point label upon fixed position
                stringBuild.Append("," + CSV.Escape(maplexOverposter.EnablePointPlacementPriorities)); //User-defined zones
                stringBuild.Append("," + CSV.Escape(maplexOverposter.PointPlacementPriorities.AboveLeft)); //Point label zone preference above left
                stringBuild.Append("," + CSV.Escape(maplexOverposter.PointPlacementPriorities.AboveCenter)); //Point label zone preference above center
                stringBuild.Append("," + CSV.Escape(maplexOverposter.PointPlacementPriorities.AboveRight)); //Point label zone preference above right
                stringBuild.Append("," + CSV.Escape(maplexOverposter.PointPlacementPriorities.CenterLeft)); //Point label zone preference center left
                stringBuild.Append("," + CSV.Escape(maplexOverposter.PointPlacementPriorities.CenterRight)); //Point label zone preference center right
                stringBuild.Append("," + CSV.Escape(maplexOverposter.PointPlacementPriorities.BelowLeft)); //Point label zone preference below left
                stringBuild.Append("," + CSV.Escape(maplexOverposter.PointPlacementPriorities.BelowCenter)); //Point label zone preference below center
                stringBuild.Append("," + CSV.Escape(maplexOverposter.PointPlacementPriorities.BelowRight)); //Point label zone preference below right

                stringBuild.Append("," + CSV.Escape(maplexOverposter.RotationProperties.Enable)); //Enable label rotation
                stringBuild.Append("," + CSV.Escape(maplexOverposter.RotationProperties.RotationField)); //Rotation field
                IMaplexRotationProperties2 maplexRotationProperties2 = maplexOverposter.RotationProperties as IMaplexRotationProperties2;
                stringBuild.Append("," + CSV.Escape(maplexRotationProperties2.AdditionalAngle)); //Additional label rotation
                stringBuild.Append("," + CSV.Escape(maplexOverposter.RotationProperties.PerpendicularToAngle)); //Perpendicular to angle
                stringBuild.Append("," + CSV.Escape(maplexOverposter.RotationProperties.RotationType)); //Rotation type
                stringBuild.Append("," + CSV.Escape(maplexRotationProperties2.AlignmentType)); //Rotation alignment type
                stringBuild.Append("," + CSV.Escape(maplexOverposter.RotationProperties.AlignLabelToAngle)); //Rotate and align label to angle

                //Placement lines
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.LineFeatureType)); //Line feature type
                stringBuild.Append("," + CSV.Escape(maplexOverposter.LinePlacementMethod)); //Line placement method
                //stringBuild.Append("," + CSV.Escape(maplexOverposter.IsStreetPlacement)); //IsStreetPlacement  //Not used at this point
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.EnableSecondaryOffset)); //Enable secondary offset
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.AllowStraddleStacking)); //Allow stacked labels to straddle lines

                stringBuild.Append("," + CSV.Escape(maplexOverposter.ConstrainOffset)); //Constrain offset
                stringBuild.Append("," + CSV.Escape(maplexOverposter.OffsetAlongLineProperties.Distance)); //Offset along line: distance
                stringBuild.Append("," + CSV.Escape(maplexOverposter.OffsetAlongLineProperties.DistanceUnit)); //Offset along line: distance unit
                stringBuild.Append("," + CSV.Escape(maplexOverposter.OffsetAlongLineProperties.LabelAnchorPoint)); //Offset along line: measure to label part
                stringBuild.Append("," + CSV.Escape(maplexOverposter.OffsetAlongLineProperties.PlacementMethod)); //Offset along line: label position
                stringBuild.Append("," + CSV.Escape(maplexOverposter.OffsetAlongLineProperties.Tolerance)); //Offset along line: tolerance
                stringBuild.Append("," + CSV.Escape(maplexOverposter.OffsetAlongLineProperties.UseLineDirection)); //Offset along line: use line direction
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.SecondaryOffsetMaximum)); //Max secondary offset
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.SecondaryOffsetMinimum)); //Min secondary offset
                stringBuild.Append("," + CSV.Escape(maplexOverposter.AlignLabelToLineDirection)); //Align label to line direction

                stringBuild.Append("," + CSV.Escape(maplexOverposter.GraticuleAlignment)); //Align label to graticule
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.GraticuleAlignmentType)); //Graticule alignment type

                //Placement lines street
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.CanPlaceLabelOnTopOfFeature)); //May place label horizontal and centered on street
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.CanReduceLeading)); //Reduce leading of stacked overrun street labels
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.CanFlipStackedStreetLabel)); //May place primary stacked name under street ending
                stringBuild.Append("," + CSV.Escape(maplexOverposter.MinimumEndOfStreetClearance)); //Min end of street clearance
                stringBuild.Append("," + CSV.Escape(maplexOverposter.PreferredEndOfStreetClearance)); //Preferred end of street clearance

                //Placement lines contours
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.ContourAlignmentType)); //Contour alignment
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.ContourMaximumAngle)); //Max contour label angle
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.ContourLadderType)); //Contour laddering

                //Placement polygons
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.PolygonFeatureType)); //Polygon feature type
                //stringBuild.Append("," + CSV.Escape(maplexOverposter.LandParcelPlacement)); //Land parcel placement //Replaced by PolygonPlacementMethod
                stringBuild.Append("," + CSV.Escape(maplexOverposter.PolygonPlacementMethod)); //Polygon placement method
                stringBuild.Append("," + CSV.Escape(maplexOverposter3.AvoidPolygonHoles)); //Avoid holes in polygons
                stringBuild.Append("," + CSV.Escape(maplexOverposter.PreferHorizontalPlacement)); //Try horizontal polygon position first
                stringBuild.Append("," + CSV.Escape(maplexOverposter.CanPlaceLabelOutsidePolygon)); //May place label outside polygon
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.EnablePolygonFixedPosition)); //Fixed position in polygon

                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonInternalZones(esriMaplexZoneIdentifier.esriMaplexAboveLeft))); //Internal polygon label zone preference above left
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonInternalZones(esriMaplexZoneIdentifier.esriMaplexAboveCenter))); //Internal polygon label zone preference above center
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonInternalZones(esriMaplexZoneIdentifier.esriMaplexAboveRight))); //Internal polygon label zone preference above right
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonInternalZones(esriMaplexZoneIdentifier.esriMaplexCenterLeft))); //Internal polygon label zone preference center left
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonInternalZones(esriMaplexZoneIdentifier.esriMaplexCenter))); //Internal polygon label zone preference center
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonInternalZones(esriMaplexZoneIdentifier.esriMaplexCenterRight))); //Internal polygon label zone preference center right
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonInternalZones(esriMaplexZoneIdentifier.esriMaplexBelowLeft))); //Internal polygon label zone preference below left
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonInternalZones(esriMaplexZoneIdentifier.esriMaplexBelowCenter))); //Internal polygon label zone preference below center
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonInternalZones(esriMaplexZoneIdentifier.esriMaplexBelowRight))); //Internal polygon label zone preference below right

                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonExternalZones(esriMaplexZoneIdentifier.esriMaplexAboveLeft))); //External polygon label zone preference above left
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonExternalZones(esriMaplexZoneIdentifier.esriMaplexAboveCenter))); //External polygon label zone preference above center
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonExternalZones(esriMaplexZoneIdentifier.esriMaplexAboveRight))); //External polygon label zone preference above right
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonExternalZones(esriMaplexZoneIdentifier.esriMaplexCenterLeft))); //External polygon label zone preference center left
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonExternalZones(esriMaplexZoneIdentifier.esriMaplexCenter))); //External polygon label zone preference center
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonExternalZones(esriMaplexZoneIdentifier.esriMaplexCenterRight))); //External polygon label zone preference center right
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonExternalZones(esriMaplexZoneIdentifier.esriMaplexBelowLeft))); //External polygon label zone preference below left
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonExternalZones(esriMaplexZoneIdentifier.esriMaplexBelowCenter))); //External polygon label zone preference below center
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonExternalZones(esriMaplexZoneIdentifier.esriMaplexBelowRight))); //External polygon label zone preference below right

                stringBuild.Append("," + CSV.Escape(maplexOverposter2.PolygonAnchorPointType)); //Polygon label anchor point

                //Placement polygon boundary
                stringBuild.Append("," + CSV.Escape(maplexOverposter3.BoundaryLabelingAllowSingleSided)); //Allow single-sided boundary labeling
                stringBuild.Append("," + CSV.Escape(maplexOverposter3.BoundaryLabelingSingleSidedOnLine)); //Center single-sided boundary labels on line
                stringBuild.Append("," + CSV.Escape(maplexOverposter3.BoundaryLabelingAllowHoles)); //Allow boundary labeling of holes

                //Placement generic
                stringBuild.Append("," + CSV.Escape(maplexOverposter.FeatureType)); //Feature type
                stringBuild.Append("," + CSV.Escape(maplexOverposter.LabelPriority)); //Label priority

                stringBuild.Append("," + CSV.Escape(maplexOverposter.PrimaryOffset)); //Primary offset
                stringBuild.Append("," + CSV.Escape(maplexOverposter.PrimaryOffsetUnit)); //Primary offset unit
                stringBuild.Append("," + CSV.Escape(maplexOverposter.SecondaryOffset)); //Maximum offset
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.IsOffsetFromFeatureGeometry)); //Measure offset from feature geometry

                stringBuild.Append("," + CSV.Escape(maplexOverposter.SpreadCharacters)); //SpreadCharacters
                stringBuild.Append("," + CSV.Escape(maplexOverposter.MaximumCharacterSpacing)); //MaximumCharacterSpacing
                stringBuild.Append("," + CSV.Escape(maplexOverposter.SpreadWords)); //SpreadWords
                stringBuild.Append("," + CSV.Escape(maplexOverposter.MaximumWordSpacing)); //MaximumWordSpacing
                //stringBuild.Append("," + CSV.Escape(maplexOverposter.FeatureBuffer)); //FeatureBuffer  //Not used at this point
                //stringBuild.Append("," + CSV.Escape(maplexOverposter.CanRemoveOverlappingLabel)); //CanRemoveOverlappingLabel  //Not used at this point

                //Fitting strategy
                stringBuild.Append("," + CSV.Escape(maplexOverposter.CanStackLabel)); //Stack
                stringBuild.Append("," + CSV.Escape(maplexOverposter.LabelStackingProperties.StackJustification)); //Label stacking alignment)
                stringBuild.Append("," + CSV.Escape(maplexOverposter.LabelStackingProperties.MaximumNumberOfLines)); //Label stacking max number of lines
                stringBuild.Append("," + CSV.Escape(maplexOverposter.LabelStackingProperties.MaximumNumberOfCharsPerLine)); //Label stacking max chars per line
                stringBuild.Append("," + CSV.Escape(maplexOverposter.LabelStackingProperties.MinimumNumberOfCharsPerLine)); //Label stacking min chars per line

                IMaplexLabelStackingProperties stackingProps = maplexOverposter.LabelStackingProperties; //Label stacking character 1-5
                for (int i = 0; i < 5; i++)
                {
                    string separatorPrint = "";
                    bool splitForced = false, Visible = false, splitAfter = false;
                    if (i < stackingProps.SeparatorCount)
                    {
                        string separator = "";
                        stackingProps.QuerySeparator(i, out separator, out Visible, out splitForced, out splitAfter);
                        separatorPrint = String.Format("U+{0:X4} ", Convert.ToUInt16(separator[0]));
                        stringBuild.Append("," + CSV.Escape(separatorPrint + "," + Visible + "," + splitForced + "," + splitAfter));
                    }
                    else //just print the remaining chars as nulls
                    {
                        separatorPrint = "null";
                        stringBuild.Append("," + CSV.Escape(separatorPrint + "," + Visible + "," + splitForced + "," + splitAfter));
                    }
                }
                stringBuild.Append("," + CSV.Escape(maplexOverposter.CanOverrunFeature)); //Overrun
                stringBuild.Append("," + CSV.Escape(maplexOverposter.MaximumLabelOverrun)); //Max label overrun
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.MaximumLabelOverrunUnit)); //Max label overrun unit
                stringBuild.Append("," + CSV.Escape(maplexOverposter.AllowAsymmetricOverrun)); //Allow asymmetric overrun
                stringBuild.Append("," + CSV.Escape(maplexOverposter.CanReduceFontSize)); //Reduce font size
                stringBuild.Append("," + CSV.Escape(maplexOverposter.FontHeightReductionLimit)); //Reduce font size lower limit
                stringBuild.Append("," + CSV.Escape(maplexOverposter.FontHeightReductionStep)); //Reduce font size by this interval
                stringBuild.Append("," + CSV.Escape(maplexOverposter.FontWidthReductionLimit)); //Compress font width lower limit
                stringBuild.Append("," + CSV.Escape(maplexOverposter.FontWidthReductionStep)); //Compress font width by this interval
                stringBuild.Append("," + CSV.Escape(maplexOverposter.CanAbbreviateLabel)); //Abbreviate
                stringBuild.Append("," + CSV.Escape(maplexOverposter.DictionaryName)); //Abbreviation dictionary
                stringBuild.Append("," + CSV.Escape(maplexOverposter.CanTruncateLabel)); //Truncate
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.TruncationMarkerCharacter)); //Truncation marker character
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.TruncationMinimumLength)); //Truncation min word length
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.TruncationPreferredCharacters)); //Truncation chars to remove
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.CanKeyNumberLabel)); //Allow key numbering
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.KeyNumberGroupName)); //Key numbering group name
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_StrategyPriority(esriMaplexStrategyIdentifier.esriMaplexStrategyStacking))); //Fitting strategy order for stacking
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_StrategyPriority(esriMaplexStrategyIdentifier.esriMaplexStrategyOverrun))); //Fitting strategy order for overrun
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_StrategyPriority(esriMaplexStrategyIdentifier.esriMaplexStrategyFontCompression))); //Fitting strategy order for font compression
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_StrategyPriority(esriMaplexStrategyIdentifier.esriMaplexStrategyFontReduction))); //Fitting strategy order for font reduction
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_StrategyPriority(esriMaplexStrategyIdentifier.esriMaplexStrategyAbbreviation))); //Fitting strategy order for abbreviation


                //Label Density
                stringBuild.Append("," + CSV.Escape(maplexOverposter.ThinDuplicateLabels)); //Remove duplicate labels
                stringBuild.Append("," + CSV.Escape(maplexOverposter.ThinningDistance)); //Remove duplicate radius
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.ThinningDistanceUnit)); //Remove duplicate radius unit
                stringBuild.Append("," + CSV.Escape(maplexOverposter.RepeatLabel)); //Repeat label
                stringBuild.Append("," + CSV.Escape(maplexOverposter.MinimumRepetitionInterval)); //Repeat label interval
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.RepetitionIntervalUnit)); //Repeat label interval unit
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.PreferLabelNearMapBorder)); //Prefer repeated label near map border
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.PreferLabelNearMapBorderClearance)); //Prefer repeated label near map border clearance
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.PreferLabelNearJunction)); //Prefer repeated label near junction
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.PreferLabelNearJunctionClearance)); //Prefer repeated label near junction clearance
                stringBuild.Append("," + CSV.Escape(maplexOverposter.LabelBuffer)); //Label buffer
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.IsLabelBufferHardConstraint)); //Label buffer hard constraint
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.IsMinimumSizeBasedOnArea)); //Min size based on area
                stringBuild.Append("," + CSV.Escape(maplexOverposter.MinimumSizeForLabeling)); //Min feature size
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.MinimumFeatureSizeUnit)); //Min feature size unit
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.ConnectionType)); //Line connection type
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.EnableConnection)); //Connect features
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.MultiPartOption)); //Unconnected line label multi-part option
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.LabelLargestPolygon)); //Label largest polygon feature part

                //Conflict Resolution
                stringBuild.Append("," + CSV.Escape(maplexOverposter.FeatureWeight)); //Feature weight
                stringBuild.Append("," + CSV.Escape(maplexOverposter.PolygonBoundaryWeight)); //Polygon boundary feature weight
                stringBuild.Append("," + CSV.Escape(maplexOverposter.BackgroundLabel)); //Background label
                stringBuild.Append("," + CSV.Escape(maplexOverposter.NeverRemoveLabel)); //Never remove

                #endregion
            }

            return stringBuild.ToString();
        }
        protected override void OnUpdate()
        {
            Enabled = ArcMap.Application != null;
        }
        private static string GetFieldColumns()
        {
            StringBuilder stringBuild = new StringBuilder();

            #region Fields
            stringBuild.Append("Map name");
            stringBuild.Append("," + "Layer name");
            stringBuild.Append("," + "Label class name");
            //stringBuild.Append("," + "Offset"); //not used at this point
            stringBuild.Append("," + "SQL query");
            stringBuild.Append("," + "Visiblity max scale");
            stringBuild.Append("," + "Visiblity min scale");
            stringBuild.Append("," + "Label expression language");
            stringBuild.Append("," + "Label expression");
            stringBuild.Append("," + "Use coded value description");
            stringBuild.Append("," + "Remove extra line breaks");
            stringBuild.Append("," + "Remove extra white space");

            //Placement points
            stringBuild.Append("," + "Point placement method");
            stringBuild.Append("," + "Measure offset from exact symbol outline");
            stringBuild.Append("," + "May shift point label upon fixed position");
            stringBuild.Append("," + "User-defined zones");
            stringBuild.Append("," + "Point label zone preference above left");
            stringBuild.Append("," + "Point label zone preference above center");
            stringBuild.Append("," + "Point label zone preference above right");
            stringBuild.Append("," + "Point label zone preference center left");
            stringBuild.Append("," + "Point label zone preference center right");
            stringBuild.Append("," + "Point label zone preference below left");
            stringBuild.Append("," + "Point label zone preference below center");
            stringBuild.Append("," + "Point label zone preference below right");
            stringBuild.Append("," + "Enable label rotation");
            stringBuild.Append("," + "Rotation field");
            stringBuild.Append("," + "Additional label rotation");
            stringBuild.Append("," + "Perpendicular to angle");
            stringBuild.Append("," + "Rotation type");
            stringBuild.Append("," + "Rotation alignment type");
            stringBuild.Append("," + "Rotate and align label to angle");

            //Placement lines
            stringBuild.Append("," + "Line feature type");
            stringBuild.Append("," + "Line placement method");
            // stringBuild.Append("," + "IsStreetPlacement"); //TODO //Replaced by LinePlacementMethod
            stringBuild.Append("," + "Enable secondary offset");
            stringBuild.Append("," + "Allow stacked labels to straddle lines");

            stringBuild.Append("," + "Constrain offset");
            stringBuild.Append("," + "Offset along line: distance");
            stringBuild.Append("," + "Offset along line: distance unit");
            stringBuild.Append("," + "Offset along line: measure to label part");
            stringBuild.Append("," + "Offset along line: label position");
            stringBuild.Append("," + "Offset along line: tolerance");
            stringBuild.Append("," + "Offset along line: use line direction");
            stringBuild.Append("," + "Max secondary offset");
            stringBuild.Append("," + "Min secondary offset");
            stringBuild.Append("," + "Align label to line direction");

            stringBuild.Append("," + "Align label to graticule");
            stringBuild.Append("," + "Graticule alignment type");

            //Placement lines street
            stringBuild.Append("," + "May place label horizontal and centered on street");
            stringBuild.Append("," + "Reduce leading of stacked overrun street labels");
            stringBuild.Append("," + "May place primary stacked name under street ending");
            stringBuild.Append("," + "Min end of street clearance");
            stringBuild.Append("," + "Preferred end of street clearance");

            //Placement lines contours
            stringBuild.Append("," + "Contour alignment");
            stringBuild.Append("," + "Max contour label angle");
            stringBuild.Append("," + "Contour laddering");

            //Placement polygons
            stringBuild.Append("," + "Polygon feature type");
            //stringBuild.Append("," + "Land parcel placement"); //Replaced by PolygonPlacementMethod
            stringBuild.Append("," + "Polygon placement method");
            stringBuild.Append("," + "Avoid holes in polygons");
            stringBuild.Append("," + "Try horizontal polygon position first");
            stringBuild.Append("," + "May place label outside polygon");
            stringBuild.Append("," + "Fixed position in polygon");

            stringBuild.Append("," + "Internal polygon label zone preference above left");
            stringBuild.Append("," + "Internal polygon label zone preference above center");
            stringBuild.Append("," + "Internal polygon label zone preference above right");
            stringBuild.Append("," + "Internal polygon label zone preference center left");
            stringBuild.Append("," + "Internal polygon label zone preference center");
            stringBuild.Append("," + "Internal polygon label zone preference center right");
            stringBuild.Append("," + "Internal polygon label zone preference below left");
            stringBuild.Append("," + "Internal polygon label zone preference below center");
            stringBuild.Append("," + "Internal polygon label zone preference below right");

            stringBuild.Append("," + "External polygon label zone preference above left");
            stringBuild.Append("," + "External polygon label zone preference above center");
            stringBuild.Append("," + "External polygon label zone preference above right");
            stringBuild.Append("," + "External polygon label zone preference center left");
            stringBuild.Append("," + "External polygon label zone preference center");
            stringBuild.Append("," + "External polygon label zone preference center right");
            stringBuild.Append("," + "External polygon label zone preference below left");
            stringBuild.Append("," + "External polygon label zone preference below center");
            stringBuild.Append("," + "External polygon label zone preference below right");

            stringBuild.Append("," + "Polygon label anchor point");

            //Placement polygon boundary
            stringBuild.Append("," + "Allow single-sided boundary labeling");
            stringBuild.Append("," + "Center single-sided boundary labels on line");
            stringBuild.Append("," + "Allow boundary labeling of holes");

            //Placement generic
            stringBuild.Append("," + "Feature type");
            stringBuild.Append("," + "Label priority");

            stringBuild.Append("," + "Primary offset");
            stringBuild.Append("," + "Primary offset unit");
            stringBuild.Append("," + "Maximum offset");
            stringBuild.Append("," + "Measure offset from feature geometry");

            stringBuild.Append("," + "Spread characters");
            stringBuild.Append("," + "Max percentage for spreading characters");
            stringBuild.Append("," + "Spread words");
            stringBuild.Append("," + "Max percentage for spreading words");
            //stringBuild.Append("," + "FeatureBuffer"); //Not used at this point
            //stringBuild.Append("," + "CanRemoveOverlappingLabel"); //Not used at this point

            //Fitting strategy
            stringBuild.Append("," + "Stack");
            stringBuild.Append("," + "Label stacking alignment");
            stringBuild.Append("," + "Label stacking max number of lines");
            stringBuild.Append("," + "Label stacking max chars per line");
            stringBuild.Append("," + "Label stacking min chars per line");

            stringBuild.Append("," + "Label stacking character 1");
            stringBuild.Append("," + "Label stacking character 2");
            stringBuild.Append("," + "Label stacking character 3");
            stringBuild.Append("," + "Label stacking character 4");
            stringBuild.Append("," + "Label stacking character 5");

            stringBuild.Append("," + "Overrun");
            stringBuild.Append("," + "Max label overrun");
            stringBuild.Append("," + "Max label overrun unit");
            stringBuild.Append("," + "Allow asymmetric overrun");
            stringBuild.Append("," + "Reduce font size");
            stringBuild.Append("," + "Reduce font size lower limit");
            stringBuild.Append("," + "Reduce font size by this interval");
            stringBuild.Append("," + "Compress font width lower limit");
            stringBuild.Append("," + "Compress font width by this interval");
            stringBuild.Append("," + "Abbreviate");
            stringBuild.Append("," + "Abbreviation dictionary");
            stringBuild.Append("," + "Truncate");
            stringBuild.Append("," + "Truncation marker character");
            stringBuild.Append("," + "Truncation min word length");
            stringBuild.Append("," + "Truncation chars to remove");
            stringBuild.Append("," + "Allow key numbering");
            stringBuild.Append("," + "Key numbering group name");
            stringBuild.Append("," + "Fitting strategy order for stacking");
            stringBuild.Append("," + "Fitting strategy order for overrun");
            stringBuild.Append("," + "Fitting strategy order for font compression");
            stringBuild.Append("," + "Fitting strategy order for font reduction");
            stringBuild.Append("," + "Fitting strategy order for abbreviation");

            //Label density
            stringBuild.Append("," + "Remove duplicate labels");
            stringBuild.Append("," + "Remove duplicates radius");
            stringBuild.Append("," + "Remove duplicates radius unit");
            stringBuild.Append("," + "Repeat label");
            stringBuild.Append("," + "Repeat label interval");
            stringBuild.Append("," + "Repeat label interval unit");
            stringBuild.Append("," + "Prefer repeated label near map border");
            stringBuild.Append("," + "Prefer repeated label near map border clearance");
            stringBuild.Append("," + "Prefer repeated label near junction");
            stringBuild.Append("," + "Prefer repeated label near junction clearance:");
            stringBuild.Append("," + "Label buffer");
            stringBuild.Append("," + "Label buffer hard constraint");
            stringBuild.Append("," + "Min size based on area");
            stringBuild.Append("," + "Min feature size");
            stringBuild.Append("," + "Min feature size unit");
            stringBuild.Append("," + "Line connection type");
            stringBuild.Append("," + "Connect features");
            stringBuild.Append("," + "Unconnected line label multi-part option");
            stringBuild.Append("," + "Label largest polygon feature part");

            //Conflict Resolution
            stringBuild.Append("," + "Feature weight");
            stringBuild.Append("," + "Polygon boundary feature weight");
            stringBuild.Append("," + "Background label");
            stringBuild.Append("," + "Never remove");

            #endregion


            return stringBuild.ToString();
        }
    }

}
