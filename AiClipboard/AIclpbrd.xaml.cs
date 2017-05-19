using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.IO;
using Corel.Interop.VGCore;
using Clipboard = System.Windows.Clipboard;
using cdrShapeType = Corel.Interop.VGCore.cdrShapeType;

namespace AiClipboard
{
    public partial class AIclpbrd : UserControl
    {
        private static Corel.Interop.VGCore.Application _dApp = null;

        public AIclpbrd() { InitializeComponent(); }
        public AIclpbrd(object app)
        {
            InitializeComponent();
            _dApp = (Corel.Interop.VGCore.Application)app;
        }

        private void CopyAi(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (_dApp.Documents.Count == 0) return;
                if (_dApp.ActiveSelectionRange.Count == 0) return;

                var tmp = Path.GetTempPath();
                var pdf = _dApp.ActiveDocument.PDFSettings;

                pdf.pdfVersion = pdfVersion.pdfVersion15; //6
                pdf.PublishRange = pdfExportRange.pdfSelection;

                pdf.ColorMode = pdfColorMode.pdfNative;
                pdf.OutputSpotColorsAs = pdfSpotType.pdfSpotAsSpot;

                pdf.TextAsCurves = false;
                pdf.TextExportMode = pdfTextExportMode.pdfTextAsUnicode;
                pdf.EmbedFonts = false;
                pdf.EmbedBaseFonts = false;
                pdf.TrueTypeToType1 = false;
                pdf.SubsetFonts = false;
                pdf.CompressText = true;

                pdf.Thumbnails = false;
                pdf.Encoding = pdfEncodingType.pdfBinary;

                pdf.BitmapCompression = pdfBitmapCompressionType.pdfLZW;
                pdf.DownsampleColor = false;
                pdf.DownsampleGray = false;
                pdf.DownsampleMono = false;
                pdf.ComplexFillsAsBitmaps = false;

                pdf.Hyperlinks = false;
                pdf.Bookmarks = false;

                pdf.Overprints = false;
                pdf.Halftones = false;
                pdf.MaintainOPILinks = false;
                pdf.FountainSteps = 256;
                pdf.EPSAs = pdfEPSAs.pdfPostscript;
                pdf.IncludeBleed = false;
                pdf.Linearize = false;
                pdf.CropMarks = false;
                pdf.RegistrationMarks = false;
                pdf.DensitometerScales = false;
                pdf.FileInformation = false;

                string pdfTempFile = tmp + _dApp.ActiveDocument.Name + "-" + _dApp.ActiveWindow.Handle.ToString(CultureInfo.InvariantCulture) + "-copy.pdf";
                _dApp.ActiveDocument.PublishToPDF(pdfTempFile);

                var ms = new MemoryStream(File.ReadAllBytes(pdfTempFile));
                var iData = new DataObject();
                iData.SetData(@"Portable Document Format", ms, false);
                Clipboard.SetDataObject(iData, false);

                File.Delete(pdfTempFile);
            }
            catch (Exception err)
            {
                MessageBox.Show(err.ToString(), "AiClipboard", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void PasteAi(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (_dApp.Documents.Count == 0) return;
                if (!_dApp.ActiveLayer.Editable || !_dApp.ActiveLayer.Visible) return;

                var iData = Clipboard.GetDataObject();
                if (iData == null) return;

                if (iData.GetDataPresent(@"Portable Document Format"))
                {
                    var tmp = Path.GetTempPath();
                    string pdfTempFile = tmp + _dApp.ActiveDocument.Name + "-" + _dApp.ActiveWindow.Handle.ToString(CultureInfo.InvariantCulture) + "-paste.pdf";

                    var ms = iData.GetData(@"Portable Document Format") as MemoryStream;
                    File.WriteAllBytes(pdfTempFile, ms.ToArray());

                    if (File.Exists(pdfTempFile))
                    {
                        BoostStart("AiClipboard: Paste");
                        Document doc = _dApp.ActiveDocument;
                        doc.ReferencePoint = cdrReferencePoint.cdrCenter;

                        Document tempDoc = _dApp.OpenDocument(pdfTempFile);
                        FixObjects(tempDoc.ActivePage.Shapes.All());
                        tempDoc.ClearSelection();
                        tempDoc.ActivePage.Shapes.All().Group().Copy();
                        
                        doc.Activate();
                        doc.ActiveLayer.Paste();
                        tempDoc.Close();

                        //var imp = new Corel.Interop.VGCore.StructImportOptions { MaintainLayers = false };
                        //dApp.ActiveLayer.ImportEx(pdfTempFile, Corel.Interop.VGCore.cdrFilter.cdrAI9, imp).Finish();

                        File.Delete(pdfTempFile);
                        BoostFinish();
                    }
                }
            }
            catch (Exception err)
            {
                MessageBox.Show(err.ToString(), "AiClipboard", MessageBoxButton.OK, MessageBoxImage.Error);
                BoostFinish();
            }
        }

        private void FixObjects(ShapeRange sr)
        {
            try
            {
                foreach (Corel.Interop.VGCore.Shape shape in sr)
                {
                    if (shape.Type == cdrShapeType.cdrGroupShape) FixObjects(shape.Shapes.All());
                    if (shape.Type == cdrShapeType.cdrSymbolShape) shape.Symbol.RevertToShapes();
                }
            }
            catch (Exception err)
            {
                MessageBox.Show(err.ToString(), "AiClipboard", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BoostStart(string undo)
        {
            _dApp.ActiveDocument.BeginCommandGroup(undo);
            _dApp.Optimization = true;
            _dApp.EventsEnabled = false;
            _dApp.ActiveDocument.SaveSettings();
            _dApp.ActiveDocument.PreserveSelection = false;
        }

        private void BoostFinish()
        {
            _dApp.ActiveDocument.PreserveSelection = true;
            _dApp.ActiveDocument.ResetSettings();
            _dApp.EventsEnabled = true;
            _dApp.Optimization = false;
            _dApp.ActiveDocument.EndCommandGroup();
            _dApp.Refresh();
        }
    }
}
