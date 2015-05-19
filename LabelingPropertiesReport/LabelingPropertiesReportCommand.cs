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

            string path = folderPath + MapDocName + ".log";
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

            stringBuild.Append(mapName); //Map name
            stringBuild.Append("," + layerName); //Layer name
            stringBuild.Append("," + annoLP.Class); //Label class name
            stringBuild.Append("," + annoLP.WhereClause); //Where clause
            stringBuild.Append("," + annoLP.AnnotationMaximumScale); //Max scale
            stringBuild.Append("," + annoLP.AnnotationMinimumScale); //Min scale
            stringBuild.Append("," + labelEngineLP.ExpressionParser.Name); //Expression Parser
            stringBuild.Append("," + labelEngineLP.Expression); //Expression
            stringBuild.Append("," + labelEngineLP.Offset); //Offset

            IMaplexOverposterLayerProperties maplexOverposter = labelEngineLP.OverposterLayerProperties as IMaplexOverposterLayerProperties;
            if (maplexOverposter != null)
            {
                #region Maplex Properties
                IMaplexOverposterLayerProperties2 maplexOverposter2 = maplexOverposter as IMaplexOverposterLayerProperties2;
                IMaplexOverposterLayerProperties3 maplexOverposter3 = maplexOverposter as IMaplexOverposterLayerProperties3;
                IMaplexOverposterLayerProperties4 maplexOverposter4 = maplexOverposter as IMaplexOverposterLayerProperties4;

                stringBuild.Append("," + maplexOverposter.AlignLabelToLineDirection); //AlignLabelToLineDirection
                stringBuild.Append("," + maplexOverposter.AllowAsymmetricOverrun); //AllowAsymmetricOverrun
                stringBuild.Append("," + maplexOverposter.BackgroundLabel); //BackgroundLabel
                stringBuild.Append("," + maplexOverposter.CanAbbreviateLabel); //CanAbbreviateLabel
                stringBuild.Append("," + maplexOverposter.CanOverrunFeature); //CanOverrunFeature
                stringBuild.Append("," + maplexOverposter.CanPlaceLabelOutsidePolygon); //CanPlaceLabelOutsidePolygon
                stringBuild.Append("," + maplexOverposter.CanReduceFontSize); //CanReduceFontSize
                stringBuild.Append("," + maplexOverposter.CanRemoveOverlappingLabel); //CanRemoveOverlappingLabel
                stringBuild.Append("," + maplexOverposter.CanShiftPointLabel); //CanShiftPointLabel
                stringBuild.Append("," + maplexOverposter.CanStackLabel); //CanStackLabel
                stringBuild.Append("," + maplexOverposter.CanTruncateLabel); //CanTruncateLabel
                stringBuild.Append("," + maplexOverposter.ConstrainOffset); //ConstrainOffset
                stringBuild.Append("," + maplexOverposter.DictionaryName); //DictionaryName
                stringBuild.Append("," + maplexOverposter.EnablePointPlacementPriorities); //EnablePointPlacementPriorities
                stringBuild.Append("," + maplexOverposter.FeatureBuffer); //FeatureBuffer
                stringBuild.Append("," + maplexOverposter.FeatureType); //FeatureType
                stringBuild.Append("," + maplexOverposter.FeatureWeight); //FeatureWeight
                stringBuild.Append("," + maplexOverposter.FontHeightReductionLimit); //FontHeightReductionLimit
                stringBuild.Append("," + maplexOverposter.FontHeightReductionStep); //FontHeightReductionStep
                stringBuild.Append("," + maplexOverposter.FontWidthReductionLimit); //FontWidthReductionLimit
                stringBuild.Append("," + maplexOverposter.FontWidthReductionStep); //FontWidthReductionStep
                stringBuild.Append("," + maplexOverposter.GraticuleAlignment); //GraticuleAlignment
                stringBuild.Append("," + maplexOverposter.IsStreetPlacement); //IsStreetPlacement
                stringBuild.Append("," + maplexOverposter.LabelBuffer); //LabelBuffer
                stringBuild.Append("," + maplexOverposter.LabelPriority); //LabelPriority
                stringBuild.Append("," + maplexOverposter.LabelStackingProperties.StackJustification); //LabelStackingProperties (Horizontal Alignment)
                stringBuild.Append("," + maplexOverposter.LabelStackingProperties.MaximumNumberOfLines); //LabelStackingProperties (Maximum Number Of Lines)
                stringBuild.Append("," + maplexOverposter.LabelStackingProperties.MaximumNumberOfCharsPerLine); //LabelStackingProperties (Maximum Number Of Chars Per Line)
                stringBuild.Append("," + maplexOverposter.LabelStackingProperties.MinimumNumberOfCharsPerLine); //LabelStackingProperties (Minimum Number Of Chars Per Line)

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
                        stringBuild.Append("," + "" + separatorPrint + "," + Visible + "," + splitForced + "," + splitAfter + "");
                    }
                    else //just print the remaining chars as nulls
                    {
                        separatorPrint = "null";
                        stringBuild.Append("," + "" + separatorPrint + "," + Visible + "," + splitForced + "," + splitAfter + "");
                    }
                }

                stringBuild.Append("," + maplexOverposter.LandParcelPlacement); //LandParcelPlacement
                stringBuild.Append("," + maplexOverposter.LinePlacementMethod); //LinePlacementMethod
                stringBuild.Append("," + maplexOverposter.MaximumCharacterSpacing); //MaximumCharacterSpacing
                stringBuild.Append("," + maplexOverposter.MaximumLabelOverrun); //MaximumLabelOverrun
                stringBuild.Append("," + maplexOverposter.MaximumWordSpacing); //MaximumWordSpacing
                stringBuild.Append("," + maplexOverposter.MinimumEndOfStreetClearance); //MinimumEndOfStreetClearance
                stringBuild.Append("," + maplexOverposter.MinimumRepetitionInterval); //MinimumRepetitionInterval
                stringBuild.Append("," + maplexOverposter.MinimumSizeForLabeling); //MinimumSizeForLabeling
                stringBuild.Append("," + maplexOverposter.NeverRemoveLabel); //NeverRemoveLabel
                stringBuild.Append("," + maplexOverposter.OffsetAlongLineProperties.Distance); //OffsetAlongLineProperties (Distance)
                stringBuild.Append("," + maplexOverposter.OffsetAlongLineProperties.DistanceUnit); //OffsetAlongLineProperties (Distance Unit)
                stringBuild.Append("," + maplexOverposter.OffsetAlongLineProperties.LabelAnchorPoint); //OffsetAlongLineProperties (Label Anchor Point)
                stringBuild.Append("," + maplexOverposter.OffsetAlongLineProperties.PlacementMethod); //OffsetAlongLineProperties (Placement Method)
                stringBuild.Append("," + maplexOverposter.OffsetAlongLineProperties.Tolerance); //OffsetAlongLineProperties (Tolerance)
                stringBuild.Append("," + maplexOverposter.OffsetAlongLineProperties.UseLineDirection); //OffsetAlongLineProperties (Use Line Direction)
                stringBuild.Append("," + maplexOverposter.PointPlacementMethod); //PointPlacementMethod
                stringBuild.Append("," + maplexOverposter.PointPlacementPriorities.AboveLeft); //PointPlacementPriorities (Above Left)
                stringBuild.Append("," + maplexOverposter.PointPlacementPriorities.AboveCenter); //PointPlacementPriorities (Above Center)
                stringBuild.Append("," + maplexOverposter.PointPlacementPriorities.AboveRight); //PointPlacementPriorities (Above Right)
                stringBuild.Append("," + maplexOverposter.PointPlacementPriorities.CenterLeft); //PointPlacementPriorities (Center Left)
                stringBuild.Append("," + maplexOverposter.PointPlacementPriorities.CenterRight); //PointPlacementPriorities (Center Right)
                stringBuild.Append("," + maplexOverposter.PointPlacementPriorities.BelowLeft); //PointPlacementPriorities (Below Left)
                stringBuild.Append("," + maplexOverposter.PointPlacementPriorities.BelowCenter); //PointPlacementPriorities (Below Center)
                stringBuild.Append("," + maplexOverposter.PointPlacementPriorities.BelowRight); //PointPlacementPriorities (Below Right)
                stringBuild.Append("," + maplexOverposter.PolygonBoundaryWeight); //PolygonBoundaryWeight
                stringBuild.Append("," + maplexOverposter.PolygonPlacementMethod); //PolygonPlacementMethod
                stringBuild.Append("," + maplexOverposter.PreferHorizontalPlacement); //PreferHorizontalPlacement
                stringBuild.Append("," + maplexOverposter.PreferredEndOfStreetClearance); //PreferredEndOfStreetClearance
                stringBuild.Append("," + maplexOverposter.PrimaryOffset); //PrimaryOffset
                stringBuild.Append("," + maplexOverposter.PrimaryOffsetUnit); //PrimaryOffsetUnit
                stringBuild.Append("," + maplexOverposter.RepeatLabel); //RepeatLabel
                stringBuild.Append("," + maplexOverposter.MinimumRepetitionInterval); //Minimum Repetition Interval
                stringBuild.Append("," + maplexOverposter.RotationProperties.AlignLabelToAngle); //RotationProperties (AlignLabelToAngle)
                stringBuild.Append("," + maplexOverposter.RotationProperties.Enable); //RotationProperties (Enable)
                stringBuild.Append("," + maplexOverposter.RotationProperties.PerpendicularToAngle); //RotationProperties (PerpendicularToAngle)
                stringBuild.Append("," + maplexOverposter.RotationProperties.RotationField); //RotationProperties (RotationField)
                stringBuild.Append("," + maplexOverposter.RotationProperties.RotationType); //RotationProperties (RotationType)
                IMaplexRotationProperties2 maplexRotationProperties2 = maplexOverposter.RotationProperties as IMaplexRotationProperties2;
                stringBuild.Append("," + maplexRotationProperties2.AdditionalAngle); //RotationProperties (AdditionalAngle)
                stringBuild.Append("," + maplexRotationProperties2.AlignmentType); //RotationProperties (AlignmentType)
                stringBuild.Append("," + maplexOverposter.SecondaryOffset); //SecondaryOffset
                stringBuild.Append("," + maplexOverposter.SpreadCharacters); //SpreadCharacters
                stringBuild.Append("," + maplexOverposter.SpreadWords); //SpreadWords
                stringBuild.Append("," + maplexOverposter.ThinDuplicateLabels); //ThinDuplicateLabels
                stringBuild.Append("," + maplexOverposter.ThinningDistance); //ThinningDistance
                stringBuild.Append("," + maplexOverposter2.CanFlipStackedStreetLabel); //CanFlipStackedStreetLabel
                stringBuild.Append("," + maplexOverposter2.CanPlaceLabelOnTopOfFeature); //CanPlaceLabelOnTopOfFeature
                stringBuild.Append("," + maplexOverposter2.CanReduceLeading); //CanReduceLeading
                stringBuild.Append("," + maplexOverposter2.ContourAlignmentType); //ContourAlignmentType
                stringBuild.Append("," + maplexOverposter2.ContourLadderType); //ContourLadderType
                stringBuild.Append("," + maplexOverposter2.ContourMaximumAngle); //ContourMaximumAngle
                stringBuild.Append("," + maplexOverposter2.EnablePolygonFixedPosition); //EnablePolygonFixedPosition
                stringBuild.Append("," + maplexOverposter2.EnableSecondaryOffset); //EnableSecondaryOffset
                stringBuild.Append("," + maplexOverposter2.GraticuleAlignmentType); //GraticuleAlignmentType
                stringBuild.Append("," + maplexOverposter2.IsLabelBufferHardConstraint); //IsLabelBufferHardConstraint
                stringBuild.Append("," + maplexOverposter2.IsMinimumSizeBasedOnArea); //IsMinimumSizeBasedOnArea
                stringBuild.Append("," + maplexOverposter2.IsOffsetFromFeatureGeometry); //IsOffsetFromFeatureGeometry
                stringBuild.Append("," + maplexOverposter2.LineFeatureType); //LineFeatureType
                stringBuild.Append("," + maplexOverposter2.MaximumLabelOverrunUnit); //MaximumLabelOverrunUnit
                stringBuild.Append("," + maplexOverposter2.MinimumFeatureSizeUnit); //MinimumFeatureSizeUnit
                stringBuild.Append("," + maplexOverposter2.PolygonAnchorPointType); //PolygonAnchorPointType
                stringBuild.Append("," + maplexOverposter2.get_PolygonExternalZones(esriMaplexZoneIdentifier.esriMaplexAboveLeft)); //PolygonExternalZones (Above Left)
                stringBuild.Append("," + maplexOverposter2.get_PolygonExternalZones(esriMaplexZoneIdentifier.esriMaplexAboveCenter)); //PolygonExternalZones (Above Center)
                stringBuild.Append("," + maplexOverposter2.get_PolygonExternalZones(esriMaplexZoneIdentifier.esriMaplexAboveRight)); //PolygonExternalZones (Above Right)
                stringBuild.Append("," + maplexOverposter2.get_PolygonExternalZones(esriMaplexZoneIdentifier.esriMaplexCenterLeft)); //PolygonExternalZones (Center Left)
                stringBuild.Append("," + maplexOverposter2.get_PolygonExternalZones(esriMaplexZoneIdentifier.esriMaplexCenter)); //PolygonExternalZones (Center)
                stringBuild.Append("," + maplexOverposter2.get_PolygonExternalZones(esriMaplexZoneIdentifier.esriMaplexCenterRight)); //PolygonExternalZones (Center Right)
                stringBuild.Append("," + maplexOverposter2.get_PolygonExternalZones(esriMaplexZoneIdentifier.esriMaplexBelowLeft)); //PolygonExternalZones (Below Left)
                stringBuild.Append("," + maplexOverposter2.get_PolygonExternalZones(esriMaplexZoneIdentifier.esriMaplexBelowCenter)); //PolygonExternalZones (Below Center)
                stringBuild.Append("," + maplexOverposter2.get_PolygonExternalZones(esriMaplexZoneIdentifier.esriMaplexBelowRight)); //PolygonExternalZones (Below Right)
                stringBuild.Append("," + maplexOverposter2.PolygonFeatureType); //PolygonFeatureType
                stringBuild.Append("," + maplexOverposter2.get_PolygonInternalZones(esriMaplexZoneIdentifier.esriMaplexAboveLeft)); //PolygonInternalZones (Above Left)
                stringBuild.Append("," + maplexOverposter2.get_PolygonInternalZones(esriMaplexZoneIdentifier.esriMaplexAboveCenter)); //PolygonInternalZones (Above Center)
                stringBuild.Append("," + maplexOverposter2.get_PolygonInternalZones(esriMaplexZoneIdentifier.esriMaplexAboveRight)); //PolygonInternalZones (Above Right)
                stringBuild.Append("," + maplexOverposter2.get_PolygonInternalZones(esriMaplexZoneIdentifier.esriMaplexCenterLeft)); //PolygonInternalZones (Center Left)
                stringBuild.Append("," + maplexOverposter2.get_PolygonInternalZones(esriMaplexZoneIdentifier.esriMaplexCenter)); //PolygonInternalZones (Center)
                stringBuild.Append("," + maplexOverposter2.get_PolygonInternalZones(esriMaplexZoneIdentifier.esriMaplexCenterRight)); //PolygonInternalZones (Center Right)
                stringBuild.Append("," + maplexOverposter2.get_PolygonInternalZones(esriMaplexZoneIdentifier.esriMaplexBelowLeft)); //PolygonInternalZones (Below Left)
                stringBuild.Append("," + maplexOverposter2.get_PolygonInternalZones(esriMaplexZoneIdentifier.esriMaplexBelowCenter)); //PolygonInternalZones (Below Center)
                stringBuild.Append("," + maplexOverposter2.get_PolygonInternalZones(esriMaplexZoneIdentifier.esriMaplexBelowRight)); //PolygonInternalZones (Below Right)
                stringBuild.Append("," + maplexOverposter2.RepetitionIntervalUnit); //RepetitionIntervalUnit
                stringBuild.Append("," + maplexOverposter2.SecondaryOffsetMaximum); //SecondaryOffsetMaximum
                stringBuild.Append("," + maplexOverposter2.SecondaryOffsetMinimum); //SecondaryOffsetMinimum
                stringBuild.Append("," + maplexOverposter2.get_StrategyPriority(esriMaplexStrategyIdentifier.esriMaplexStrategyStacking)); //StrategyPriority (Stacking)
                stringBuild.Append("," + maplexOverposter2.get_StrategyPriority(esriMaplexStrategyIdentifier.esriMaplexStrategyOverrun)); //StrategyPriority (Overrun)
                stringBuild.Append("," + maplexOverposter2.get_StrategyPriority(esriMaplexStrategyIdentifier.esriMaplexStrategyFontCompression)); //StrategyPriority (Font Compression)
                stringBuild.Append("," + maplexOverposter2.get_StrategyPriority(esriMaplexStrategyIdentifier.esriMaplexStrategyFontReduction)); //StrategyPriority (Font Reduction)
                stringBuild.Append("," + maplexOverposter2.get_StrategyPriority(esriMaplexStrategyIdentifier.esriMaplexStrategyAbbreviation)); //StrategyPriority (Abbreviation)
                stringBuild.Append("," + maplexOverposter2.ThinningDistanceUnit); //ThinningDistanceUnit
                stringBuild.Append("," + maplexOverposter3.AvoidPolygonHoles); //AvoidPolygonHoles
                stringBuild.Append("," + maplexOverposter3.BoundaryLabelingAllowHoles); //BoundaryLabelingAllowHoles
                stringBuild.Append("," + maplexOverposter3.BoundaryLabelingAllowSingleSided); //BoundaryLabelingAllowSingleSided
                stringBuild.Append("," + maplexOverposter3.BoundaryLabelingSingleSidedOnLine); //BoundaryLabelingSingleSidedOnLine
                stringBuild.Append("," + maplexOverposter4.AllowStraddleStacking); //AllowStraddleStacking
                stringBuild.Append("," + maplexOverposter4.CanKeyNumberLabel); //CanKeyNumberLabel
                stringBuild.Append("," + maplexOverposter4.ConnectionType); //ConnectionType
                stringBuild.Append("," + maplexOverposter4.EnableConnection); //EnableConnection
                stringBuild.Append("," + maplexOverposter4.KeyNumberGroupName); //KeyNumberGroupName
                stringBuild.Append("," + maplexOverposter4.LabelLargestPolygon); //LabelLargestPolygon
                stringBuild.Append("," + maplexOverposter4.MultiPartOption); //MultiPartOption
                stringBuild.Append("," + maplexOverposter4.PreferLabelNearJunction); //PreferLabelNearJunction
                stringBuild.Append("," + maplexOverposter4.PreferLabelNearJunctionClearance); //PreferLabelNearJunctionClearance
                stringBuild.Append("," + maplexOverposter4.PreferLabelNearMapBorder); //PreferLabelNearMapBorder
                stringBuild.Append("," + maplexOverposter4.PreferLabelNearMapBorderClearance); //PreferLabelNearMapBorderClearance
                stringBuild.Append("," + maplexOverposter4.RemoveExtraLineBreaks); //RemoveExtraLineBreaks
                stringBuild.Append("," + maplexOverposter4.RemoveExtraWhiteSpace); //RemoveExtraWhiteSpace
                stringBuild.Append("," + maplexOverposter4.TruncationMarkerCharacter); //TruncationMarkerCharacter
                stringBuild.Append("," + maplexOverposter4.TruncationMinimumLength); //TruncationMinimumLength
                stringBuild.Append("," + maplexOverposter4.TruncationPreferredCharacters); //TruncationPreferredCharacters
                stringBuild.Append("," + maplexOverposter4.UseExactSymbolOutline); //UseExactSymbolOutline
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
