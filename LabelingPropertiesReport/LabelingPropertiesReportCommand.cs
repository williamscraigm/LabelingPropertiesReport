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
            stringBuild.Append("," + CSV.Escape(annoLP.WhereClause)); //Where clause
            stringBuild.Append("," + CSV.Escape(annoLP.AnnotationMaximumScale)); //Max scale
            stringBuild.Append("," + CSV.Escape(annoLP.AnnotationMinimumScale)); //Min scale
            stringBuild.Append("," + CSV.Escape(labelEngineLP.ExpressionParser.Name)); //Expression Parser
            stringBuild.Append("," + CSV.Escape(labelEngineLP.Expression)); //Expression
            stringBuild.Append("," + CSV.Escape(labelEngineLP.Offset)); //Offset

            IMaplexOverposterLayerProperties maplexOverposter = labelEngineLP.OverposterLayerProperties as IMaplexOverposterLayerProperties;
            if (maplexOverposter != null)
            {
                #region Maplex Properties
                IMaplexOverposterLayerProperties2 maplexOverposter2 = maplexOverposter as IMaplexOverposterLayerProperties2;
                IMaplexOverposterLayerProperties3 maplexOverposter3 = maplexOverposter as IMaplexOverposterLayerProperties3;
                IMaplexOverposterLayerProperties4 maplexOverposter4 = maplexOverposter as IMaplexOverposterLayerProperties4;

                stringBuild.Append("," + CSV.Escape(maplexOverposter.AlignLabelToLineDirection)); //AlignLabelToLineDirection
                stringBuild.Append("," + CSV.Escape(maplexOverposter.AllowAsymmetricOverrun)); //AllowAsymmetricOverrun
                stringBuild.Append("," + CSV.Escape(maplexOverposter.BackgroundLabel)); //BackgroundLabel
                stringBuild.Append("," + CSV.Escape(maplexOverposter.CanAbbreviateLabel)); //CanAbbreviateLabel
                stringBuild.Append("," + CSV.Escape(maplexOverposter.CanOverrunFeature)); //CanOverrunFeature
                stringBuild.Append("," + CSV.Escape(maplexOverposter.CanPlaceLabelOutsidePolygon)); //CanPlaceLabelOutsidePolygon
                stringBuild.Append("," + CSV.Escape(maplexOverposter.CanReduceFontSize)); //CanReduceFontSize
                stringBuild.Append("," + CSV.Escape(maplexOverposter.CanRemoveOverlappingLabel)); //CanRemoveOverlappingLabel
                stringBuild.Append("," + CSV.Escape(maplexOverposter.CanShiftPointLabel)); //CanShiftPointLabel
                stringBuild.Append("," + CSV.Escape(maplexOverposter.CanStackLabel)); //CanStackLabel
                stringBuild.Append("," + CSV.Escape(maplexOverposter.CanTruncateLabel)); //CanTruncateLabel
                stringBuild.Append("," + CSV.Escape(maplexOverposter.ConstrainOffset)); //ConstrainOffset
                stringBuild.Append("," + CSV.Escape(maplexOverposter.DictionaryName)); //DictionaryName
                stringBuild.Append("," + CSV.Escape(maplexOverposter.EnablePointPlacementPriorities)); //EnablePointPlacementPriorities
                stringBuild.Append("," + CSV.Escape(maplexOverposter.FeatureBuffer)); //FeatureBuffer
                stringBuild.Append("," + CSV.Escape(maplexOverposter.FeatureType)); //FeatureType
                stringBuild.Append("," + CSV.Escape(maplexOverposter.FeatureWeight)); //FeatureWeight
                stringBuild.Append("," + CSV.Escape(maplexOverposter.FontHeightReductionLimit)); //FontHeightReductionLimit
                stringBuild.Append("," + CSV.Escape(maplexOverposter.FontHeightReductionStep)); //FontHeightReductionStep
                stringBuild.Append("," + CSV.Escape(maplexOverposter.FontWidthReductionLimit)); //FontWidthReductionLimit
                stringBuild.Append("," + CSV.Escape(maplexOverposter.FontWidthReductionStep)); //FontWidthReductionStep
                stringBuild.Append("," + CSV.Escape(maplexOverposter.GraticuleAlignment)); //GraticuleAlignment
                stringBuild.Append("," + CSV.Escape(maplexOverposter.IsStreetPlacement)); //IsStreetPlacement
                stringBuild.Append("," + CSV.Escape(maplexOverposter.LabelBuffer)); //LabelBuffer
                stringBuild.Append("," + CSV.Escape(maplexOverposter.LabelPriority)); //LabelPriority
                stringBuild.Append("," + CSV.Escape(maplexOverposter.LabelStackingProperties.StackJustification)); //LabelStackingProperties (Horizontal Alignment)
                stringBuild.Append("," + CSV.Escape(maplexOverposter.LabelStackingProperties.MaximumNumberOfLines)); //LabelStackingProperties (Maximum Number Of Lines)
                stringBuild.Append("," + CSV.Escape(maplexOverposter.LabelStackingProperties.MaximumNumberOfCharsPerLine)); //LabelStackingProperties (Maximum Number Of Chars Per Line)
                stringBuild.Append("," + CSV.Escape(maplexOverposter.LabelStackingProperties.MinimumNumberOfCharsPerLine)); //LabelStackingProperties (Minimum Number Of Chars Per Line)

                IMaplexLabelStackingProperties stackingProps = maplexOverposter.LabelStackingProperties;
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

                stringBuild.Append("," + CSV.Escape(maplexOverposter.LandParcelPlacement)); //LandParcelPlacement
                stringBuild.Append("," + CSV.Escape(maplexOverposter.LinePlacementMethod)); //LinePlacementMethod
                stringBuild.Append("," + CSV.Escape(maplexOverposter.MaximumCharacterSpacing)); //MaximumCharacterSpacing
                stringBuild.Append("," + CSV.Escape(maplexOverposter.MaximumLabelOverrun)); //MaximumLabelOverrun
                stringBuild.Append("," + CSV.Escape(maplexOverposter.MaximumWordSpacing)); //MaximumWordSpacing
                stringBuild.Append("," + CSV.Escape(maplexOverposter.MinimumEndOfStreetClearance)); //MinimumEndOfStreetClearance
                stringBuild.Append("," + CSV.Escape(maplexOverposter.MinimumRepetitionInterval)); //MinimumRepetitionInterval
                stringBuild.Append("," + CSV.Escape(maplexOverposter.MinimumSizeForLabeling)); //MinimumSizeForLabeling
                stringBuild.Append("," + CSV.Escape(maplexOverposter.NeverRemoveLabel)); //NeverRemoveLabel
                stringBuild.Append("," + CSV.Escape(maplexOverposter.OffsetAlongLineProperties.Distance)); //OffsetAlongLineProperties (Distance)
                stringBuild.Append("," + CSV.Escape(maplexOverposter.OffsetAlongLineProperties.DistanceUnit)); //OffsetAlongLineProperties (Distance Unit)
                stringBuild.Append("," + CSV.Escape(maplexOverposter.OffsetAlongLineProperties.LabelAnchorPoint)); //OffsetAlongLineProperties (Label Anchor Point)
                stringBuild.Append("," + CSV.Escape(maplexOverposter.OffsetAlongLineProperties.PlacementMethod)); //OffsetAlongLineProperties (Placement Method)
                stringBuild.Append("," + CSV.Escape(maplexOverposter.OffsetAlongLineProperties.Tolerance)); //OffsetAlongLineProperties (Tolerance)
                stringBuild.Append("," + CSV.Escape(maplexOverposter.OffsetAlongLineProperties.UseLineDirection)); //OffsetAlongLineProperties (Use Line Direction)
                stringBuild.Append("," + CSV.Escape(maplexOverposter.PointPlacementMethod)); //PointPlacementMethod
                stringBuild.Append("," + CSV.Escape(maplexOverposter.PointPlacementPriorities.AboveLeft)); //PointPlacementPriorities (Above Left)
                stringBuild.Append("," + CSV.Escape(maplexOverposter.PointPlacementPriorities.AboveCenter)); //PointPlacementPriorities (Above Center)
                stringBuild.Append("," + CSV.Escape(maplexOverposter.PointPlacementPriorities.AboveRight)); //PointPlacementPriorities (Above Right)
                stringBuild.Append("," + CSV.Escape(maplexOverposter.PointPlacementPriorities.CenterLeft)); //PointPlacementPriorities (Center Left)
                stringBuild.Append("," + CSV.Escape(maplexOverposter.PointPlacementPriorities.CenterRight)); //PointPlacementPriorities (Center Right)
                stringBuild.Append("," + CSV.Escape(maplexOverposter.PointPlacementPriorities.BelowLeft)); //PointPlacementPriorities (Below Left)
                stringBuild.Append("," + CSV.Escape(maplexOverposter.PointPlacementPriorities.BelowCenter)); //PointPlacementPriorities (Below Center)
                stringBuild.Append("," + CSV.Escape(maplexOverposter.PointPlacementPriorities.BelowRight)); //PointPlacementPriorities (Below Right)
                stringBuild.Append("," + CSV.Escape(maplexOverposter.PolygonBoundaryWeight)); //PolygonBoundaryWeight
                stringBuild.Append("," + CSV.Escape(maplexOverposter.PolygonPlacementMethod)); //PolygonPlacementMethod
                stringBuild.Append("," + CSV.Escape(maplexOverposter.PreferHorizontalPlacement)); //PreferHorizontalPlacement
                stringBuild.Append("," + CSV.Escape(maplexOverposter.PreferredEndOfStreetClearance)); //PreferredEndOfStreetClearance
                stringBuild.Append("," + CSV.Escape(maplexOverposter.PrimaryOffset)); //PrimaryOffset
                stringBuild.Append("," + CSV.Escape(maplexOverposter.PrimaryOffsetUnit)); //PrimaryOffsetUnit
                stringBuild.Append("," + CSV.Escape(maplexOverposter.RepeatLabel)); //RepeatLabel
                stringBuild.Append("," + CSV.Escape(maplexOverposter.MinimumRepetitionInterval)); //Minimum Repetition Interval
                stringBuild.Append("," + CSV.Escape(maplexOverposter.RotationProperties.AlignLabelToAngle)); //RotationProperties (AlignLabelToAngle)
                stringBuild.Append("," + CSV.Escape(maplexOverposter.RotationProperties.Enable)); //RotationProperties (Enable)
                stringBuild.Append("," + CSV.Escape(maplexOverposter.RotationProperties.PerpendicularToAngle)); //RotationProperties (PerpendicularToAngle)
                stringBuild.Append("," + CSV.Escape(maplexOverposter.RotationProperties.RotationField)); //RotationProperties (RotationField)
                stringBuild.Append("," + CSV.Escape(maplexOverposter.RotationProperties.RotationType)); //RotationProperties (RotationType)
                IMaplexRotationProperties2 maplexRotationProperties2 = maplexOverposter.RotationProperties as IMaplexRotationProperties2;
                stringBuild.Append("," + CSV.Escape(maplexRotationProperties2.AdditionalAngle)); //RotationProperties (AdditionalAngle)
                stringBuild.Append("," + CSV.Escape(maplexRotationProperties2.AlignmentType)); //RotationProperties (AlignmentType)
                stringBuild.Append("," + CSV.Escape(maplexOverposter.SecondaryOffset)); //SecondaryOffset
                stringBuild.Append("," + CSV.Escape(maplexOverposter.SpreadCharacters)); //SpreadCharacters
                stringBuild.Append("," + CSV.Escape(maplexOverposter.SpreadWords)); //SpreadWords
                stringBuild.Append("," + CSV.Escape(maplexOverposter.ThinDuplicateLabels)); //ThinDuplicateLabels
                stringBuild.Append("," + CSV.Escape(maplexOverposter.ThinningDistance)); //ThinningDistance
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.CanFlipStackedStreetLabel)); //CanFlipStackedStreetLabel
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.CanPlaceLabelOnTopOfFeature)); //CanPlaceLabelOnTopOfFeature
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.CanReduceLeading)); //CanReduceLeading
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.ContourAlignmentType)); //ContourAlignmentType
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.ContourLadderType)); //ContourLadderType
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.ContourMaximumAngle)); //ContourMaximumAngle
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.EnablePolygonFixedPosition)); //EnablePolygonFixedPosition
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.EnableSecondaryOffset)); //EnableSecondaryOffset
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.GraticuleAlignmentType)); //GraticuleAlignmentType
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.IsLabelBufferHardConstraint)); //IsLabelBufferHardConstraint
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.IsMinimumSizeBasedOnArea)); //IsMinimumSizeBasedOnArea
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.IsOffsetFromFeatureGeometry)); //IsOffsetFromFeatureGeometry
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.LineFeatureType)); //LineFeatureType
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.MaximumLabelOverrunUnit)); //MaximumLabelOverrunUnit
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.MinimumFeatureSizeUnit)); //MinimumFeatureSizeUnit
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.PolygonAnchorPointType)); //PolygonAnchorPointType
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonExternalZones(esriMaplexZoneIdentifier.esriMaplexAboveLeft))); //PolygonExternalZones (Above Left)
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonExternalZones(esriMaplexZoneIdentifier.esriMaplexAboveCenter))); //PolygonExternalZones (Above Center)
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonExternalZones(esriMaplexZoneIdentifier.esriMaplexAboveRight))); //PolygonExternalZones (Above Right)
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonExternalZones(esriMaplexZoneIdentifier.esriMaplexCenterLeft))); //PolygonExternalZones (Center Left)
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonExternalZones(esriMaplexZoneIdentifier.esriMaplexCenter))); //PolygonExternalZones (Center)
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonExternalZones(esriMaplexZoneIdentifier.esriMaplexCenterRight))); //PolygonExternalZones (Center Right)
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonExternalZones(esriMaplexZoneIdentifier.esriMaplexBelowLeft))); //PolygonExternalZones (Below Left)
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonExternalZones(esriMaplexZoneIdentifier.esriMaplexBelowCenter))); //PolygonExternalZones (Below Center)
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonExternalZones(esriMaplexZoneIdentifier.esriMaplexBelowRight))); //PolygonExternalZones (Below Right)
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.PolygonFeatureType)); //PolygonFeatureType
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonInternalZones(esriMaplexZoneIdentifier.esriMaplexAboveLeft))); //PolygonInternalZones (Above Left)
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonInternalZones(esriMaplexZoneIdentifier.esriMaplexAboveCenter))); //PolygonInternalZones (Above Center)
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonInternalZones(esriMaplexZoneIdentifier.esriMaplexAboveRight))); //PolygonInternalZones (Above Right)
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonInternalZones(esriMaplexZoneIdentifier.esriMaplexCenterLeft))); //PolygonInternalZones (Center Left)
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonInternalZones(esriMaplexZoneIdentifier.esriMaplexCenter))); //PolygonInternalZones (Center)
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonInternalZones(esriMaplexZoneIdentifier.esriMaplexCenterRight))); //PolygonInternalZones (Center Right)
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonInternalZones(esriMaplexZoneIdentifier.esriMaplexBelowLeft))); //PolygonInternalZones (Below Left)
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonInternalZones(esriMaplexZoneIdentifier.esriMaplexBelowCenter))); //PolygonInternalZones (Below Center)
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_PolygonInternalZones(esriMaplexZoneIdentifier.esriMaplexBelowRight))); //PolygonInternalZones (Below Right)
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.RepetitionIntervalUnit)); //RepetitionIntervalUnit
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.SecondaryOffsetMaximum)); //SecondaryOffsetMaximum
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.SecondaryOffsetMinimum)); //SecondaryOffsetMinimum
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_StrategyPriority(esriMaplexStrategyIdentifier.esriMaplexStrategyStacking))); //StrategyPriority (Stacking)
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_StrategyPriority(esriMaplexStrategyIdentifier.esriMaplexStrategyOverrun))); //StrategyPriority (Overrun)
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_StrategyPriority(esriMaplexStrategyIdentifier.esriMaplexStrategyFontCompression))); //StrategyPriority (Font Compression)
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_StrategyPriority(esriMaplexStrategyIdentifier.esriMaplexStrategyFontReduction))); //StrategyPriority (Font Reduction)
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.get_StrategyPriority(esriMaplexStrategyIdentifier.esriMaplexStrategyAbbreviation))); //StrategyPriority (Abbreviation)
                stringBuild.Append("," + CSV.Escape(maplexOverposter2.ThinningDistanceUnit)); //ThinningDistanceUnit
                stringBuild.Append("," + CSV.Escape(maplexOverposter3.AvoidPolygonHoles)); //AvoidPolygonHoles
                stringBuild.Append("," + CSV.Escape(maplexOverposter3.BoundaryLabelingAllowHoles)); //BoundaryLabelingAllowHoles
                stringBuild.Append("," + CSV.Escape(maplexOverposter3.BoundaryLabelingAllowSingleSided)); //BoundaryLabelingAllowSingleSided
                stringBuild.Append("," + CSV.Escape(maplexOverposter3.BoundaryLabelingSingleSidedOnLine)); //BoundaryLabelingSingleSidedOnLine
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.AllowStraddleStacking)); //AllowStraddleStacking
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.CanKeyNumberLabel)); //CanKeyNumberLabel
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.ConnectionType)); //ConnectionType
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.EnableConnection)); //EnableConnection
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.KeyNumberGroupName)); //KeyNumberGroupName
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.LabelLargestPolygon)); //LabelLargestPolygon
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.MultiPartOption)); //MultiPartOption
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.PreferLabelNearJunction)); //PreferLabelNearJunction
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.PreferLabelNearJunctionClearance)); //PreferLabelNearJunctionClearance
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.PreferLabelNearMapBorder)); //PreferLabelNearMapBorder
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.PreferLabelNearMapBorderClearance)); //PreferLabelNearMapBorderClearance
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.RemoveExtraLineBreaks)); //RemoveExtraLineBreaks
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.RemoveExtraWhiteSpace)); //RemoveExtraWhiteSpace
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.TruncationMarkerCharacter)); //TruncationMarkerCharacter
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.TruncationMinimumLength)); //TruncationMinimumLength
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.TruncationPreferredCharacters)); //TruncationPreferredCharacters
                stringBuild.Append("," + CSV.Escape(maplexOverposter4.UseExactSymbolOutline)); //UseExactSymbolOutline
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
            stringBuild.Append("," + "Where clause");
            stringBuild.Append("," + "Max scale");
            stringBuild.Append("," + "Min scale");
            stringBuild.Append("," + "Expression Parser");
            stringBuild.Append("," + "Expression");
            stringBuild.Append("," + "Offset");

            stringBuild.Append("," + "AlignLabelToLineDirection");
            stringBuild.Append("," + "AllowAsymmetricOverrun");
            stringBuild.Append("," + "BackgroundLabel");
            stringBuild.Append("," + "CanAbbreviateLabel");
            stringBuild.Append("," + "CanOverrunFeature");
            stringBuild.Append("," + "CanPlaceLabelOutsidePolygon");
            stringBuild.Append("," + "CanReduceFontSize");
            stringBuild.Append("," + "CanRemoveOverlappingLabel");
            stringBuild.Append("," + "CanShiftPointLabel");
            stringBuild.Append("," + "CanStackLabel");
            stringBuild.Append("," + "CanTruncateLabel");
            stringBuild.Append("," + "ConstrainOffset");
            stringBuild.Append("," + "DictionaryName");
            stringBuild.Append("," + "EnablePointPlacementPriorities");
            stringBuild.Append("," + "FeatureBuffer");
            stringBuild.Append("," + "FeatureType");
            stringBuild.Append("," + "FeatureWeight");
            stringBuild.Append("," + "FontHeightReductionLimit");
            stringBuild.Append("," + "FontHeightReductionStep");
            stringBuild.Append("," + "FontWidthReductionLimit");
            stringBuild.Append("," + "FontWidthReductionStep");
            stringBuild.Append("," + "GraticuleAlignment");
            stringBuild.Append("," + "IsStreetPlacement");
            stringBuild.Append("," + "LabelBuffer");
            stringBuild.Append("," + "LabelPriority");
            stringBuild.Append("," + "LabelStackingProperties (Horizontal Alignment)");
            stringBuild.Append("," + "LabelStackingProperties (Maximum Number Of Lines)");
            stringBuild.Append("," + "LabelStackingProperties (Maximum Number Of Chars Per Line)");
            stringBuild.Append("," + "LabelStackingProperties (Minimum Number Of Chars Per Line)");

            stringBuild.Append("," + "LabelStackingProperties 1");
            stringBuild.Append("," + "LabelStackingProperties 2");
            stringBuild.Append("," + "LabelStackingProperties 3");
            stringBuild.Append("," + "LabelStackingProperties 4");
            stringBuild.Append("," + "LabelStackingProperties 5");

            stringBuild.Append("," + "LandParcelPlacement");
            stringBuild.Append("," + "LinePlacementMethod");
            stringBuild.Append("," + "MaximumCharacterSpacing");
            stringBuild.Append("," + "MaximumLabelOverrun");
            stringBuild.Append("," + "MaximumWordSpacing");
            stringBuild.Append("," + "MinimumEndOfStreetClearance");
            stringBuild.Append("," + "MinimumRepetitionInterval");
            stringBuild.Append("," + "MinimumSizeForLabeling");
            stringBuild.Append("," + "NeverRemoveLabel");
            stringBuild.Append("," + "OffsetAlongLineProperties (Distance)");
            stringBuild.Append("," + "OffsetAlongLineProperties (Distance Unit)");
            stringBuild.Append("," + "OffsetAlongLineProperties (Label Anchor Point)");
            stringBuild.Append("," + "OffsetAlongLineProperties (Placement Method)");
            stringBuild.Append("," + "OffsetAlongLineProperties (Tolerance)");
            stringBuild.Append("," + "OffsetAlongLineProperties (Use Line Direction)");
            stringBuild.Append("," + "PointPlacementMethod");
            stringBuild.Append("," + "PointPlacementPriorities (Above Left)");
            stringBuild.Append("," + "PointPlacementPriorities (Above Center)");
            stringBuild.Append("," + "PointPlacementPriorities (Above Right)");
            stringBuild.Append("," + "PointPlacementPriorities (Center Left)");
            stringBuild.Append("," + "PointPlacementPriorities (Center Right)");
            stringBuild.Append("," + "PointPlacementPriorities (Below Left)");
            stringBuild.Append("," + "PointPlacementPriorities (Below Center)");
            stringBuild.Append("," + "PointPlacementPriorities (Below Right)");
            stringBuild.Append("," + "PolygonBoundaryWeight");
            stringBuild.Append("," + "PolygonPlacementMethod");
            stringBuild.Append("," + "PreferHorizontalPlacement");
            stringBuild.Append("," + "PreferredEndOfStreetClearance");
            stringBuild.Append("," + "PrimaryOffset");
            stringBuild.Append("," + "PrimaryOffsetUnit");
            stringBuild.Append("," + "RepeatLabel");
            stringBuild.Append("," + "Minimum Repetition Interval");
            stringBuild.Append("," + "RotationProperties (AlignLabelToAngle)");
            stringBuild.Append("," + "RotationProperties (Enable)");
            stringBuild.Append("," + "RotationProperties (PerpendicularToAngle)");
            stringBuild.Append("," + "RotationProperties (RotationField)");
            stringBuild.Append("," + "RotationProperties (RotationType)");

            stringBuild.Append("," + "RotationProperties (AdditionalAngle)");
            stringBuild.Append("," + "RotationProperties (AlignmentType)");
            stringBuild.Append("," + "SecondaryOffset");
            stringBuild.Append("," + "SpreadCharacters");
            stringBuild.Append("," + "SpreadWords");
            stringBuild.Append("," + "ThinDuplicateLabels");
            stringBuild.Append("," + "ThinningDistance");
            stringBuild.Append("," + "CanFlipStackedStreetLabel");
            stringBuild.Append("," + "CanPlaceLabelOnTopOfFeature");
            stringBuild.Append("," + "CanReduceLeading");
            stringBuild.Append("," + "ContourAlignmentType");
            stringBuild.Append("," + "ContourLadderType");
            stringBuild.Append("," + "ContourMaximumAngle");
            stringBuild.Append("," + "EnablePolygonFixedPosition");
            stringBuild.Append("," + "EnableSecondaryOffset");
            stringBuild.Append("," + "GraticuleAlignmentType");
            stringBuild.Append("," + "IsLabelBufferHardConstraint");
            stringBuild.Append("," + "IsMinimumSizeBasedOnArea");
            stringBuild.Append("," + "IsOffsetFromFeatureGeometry");
            stringBuild.Append("," + "LineFeatureType");
            stringBuild.Append("," + "MaximumLabelOverrunUnit");
            stringBuild.Append("," + "MinimumFeatureSizeUnit");
            stringBuild.Append("," + "PolygonAnchorPointType");
            stringBuild.Append("," + "PolygonExternalZones (Above Left)");
            stringBuild.Append("," + "PolygonExternalZones (Above Center)");
            stringBuild.Append("," + "PolygonExternalZones (Above Right)");
            stringBuild.Append("," + "PolygonExternalZones (Center Left)");
            stringBuild.Append("," + "PolygonExternalZones (Center)");
            stringBuild.Append("," + "PolygonExternalZones (Center Right)");
            stringBuild.Append("," + "PolygonExternalZones (Below Left)");
            stringBuild.Append("," + "PolygonExternalZones (Below Center)");
            stringBuild.Append("," + "PolygonExternalZones (Below Right)");
            stringBuild.Append("," + "PolygonFeatureType");
            stringBuild.Append("," + "PolygonInternalZones (Above Left)");
            stringBuild.Append("," + "PolygonInternalZones (Above Center)");
            stringBuild.Append("," + "PolygonInternalZones (Above Right)");
            stringBuild.Append("," + "PolygonInternalZones (Center Left)");
            stringBuild.Append("," + "PolygonInternalZones (Center)");
            stringBuild.Append("," + "PolygonInternalZones (Center Right)");
            stringBuild.Append("," + "PolygonInternalZones (Below Left)");
            stringBuild.Append("," + "PolygonInternalZones (Below Center)");
            stringBuild.Append("," + "PolygonInternalZones (Below Right)");
            stringBuild.Append("," + "RepetitionIntervalUnit");
            stringBuild.Append("," + "SecondaryOffsetMaximum");
            stringBuild.Append("," + "SecondaryOffsetMinimum");
            stringBuild.Append("," + "StrategyPriority (Stacking)");
            stringBuild.Append("," + "StrategyPriority (Overrun)");
            stringBuild.Append("," + "StrategyPriority (Font Compression)");
            stringBuild.Append("," + "StrategyPriority (Font Reduction)");
            stringBuild.Append("," + "StrategyPriority (Abbreviation)");
            stringBuild.Append("," + "ThinningDistanceUnit");
            stringBuild.Append("," + "AvoidPolygonHoles");
            stringBuild.Append("," + "BoundaryLabelingAllowHoles");
            stringBuild.Append("," + "BoundaryLabelingAllowSingleSided");
            stringBuild.Append("," + "BoundaryLabelingSingleSidedOnLine");
            stringBuild.Append("," + "AllowStraddleStacking");
            stringBuild.Append("," + "CanKeyNumberLabel");
            stringBuild.Append("," + "ConnectionType");
            stringBuild.Append("," + "EnableConnection");
            stringBuild.Append("," + "KeyNumberGroupName");
            stringBuild.Append("," + "LabelLargestPolygon");
            stringBuild.Append("," + "MultiPartOption");
            stringBuild.Append("," + "PreferLabelNearJunction");
            stringBuild.Append("," + "PreferLabelNearJunctionClearance:");
            stringBuild.Append("," + "PreferLabelNearMapBorder");
            stringBuild.Append("," + "PreferLabelNearMapBorderClearance");
            stringBuild.Append("," + "RemoveExtraLineBreaks");
            stringBuild.Append("," + "RemoveExtraWhiteSpace");
            stringBuild.Append("," + "TruncationMarkerCharacter");
            stringBuild.Append("," + "TruncationMinimumLength");
            stringBuild.Append("," + "TruncationPreferredCharacters");
            stringBuild.Append("," + "UseExactSymbolOutline");
            #endregion


            return stringBuild.ToString();
        }
    }

}
