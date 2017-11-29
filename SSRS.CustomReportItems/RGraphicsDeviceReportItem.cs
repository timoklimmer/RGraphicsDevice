using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Microsoft.ReportingServices.OnDemandReportRendering;
using Image = System.Drawing.Image;

// note: do not remove this namespace declaration as this will for some reasons break the preview in VS
//       (although this is already exactly the default namespace set in the project properties)

namespace SSRS.CustomReportItems
{
    /// <summary>
    ///     Shows the graphics device output of an arbitrary R code in Reporting Services.
    /// </summary>
    public class RGraphicsDeviceReportItem : ICustomReportItem
    {
        /// <summary>
        ///     Implements the GenerateReportItemDefinition() method from interface <see cref="ICustomReportItem" />.
        /// </summary>
        public void GenerateReportItemDefinition(CustomReportItem customReportItem)
        {
            // generate the image definition object being used later to render the output image
            customReportItem.CreateCriImageDefinition();

            // get the generated report item and the related ReportingServices.OnDemandReportRendering.Image
            var generatedReportItem = customReportItem.GeneratedReportItem;
            var onDemandReportRenderingImage = (ReportingServices.OnDemandReportRendering.Image) generatedReportItem;

            // adopt the border settings
            onDemandReportRenderingImage.Style.Border.Instance.Color = customReportItem.Style.Border.Instance.Color;
            onDemandReportRenderingImage.Style.Border.Instance.Style = customReportItem.Style.Border.Instance.Style;
            onDemandReportRenderingImage.Style.Border.Instance.Width = customReportItem.Style.Border.Instance.Width;

            // set the SSRS image object's sizing to FitProportional to ensure that the image size is always right
            // note: due to resolution mismatch of the output device and our generated image, we could get wrong image sizes
            //       in the final output. the FitProportional setting will avoid this and ensure that our image has always
            //       the desired size, independent of its resolution.
            onDemandReportRenderingImage.ImageInstance.MIMEType = "image/png";
            onDemandReportRenderingImage.Sizing =
                ReportingServices.OnDemandReportRendering.Image.Sizings.FitProportional;
        }

        /// <summary>
        ///     Implements the EvaluateReportItemInstance() method from interface <see cref="ICustomReportItem" />.
        /// </summary>
        public void EvaluateReportItemInstance(CustomReportItem customReportItem)
        {
            // initialization
            var onDemandReportRenderingImage =
                ((ReportingServices.OnDemandReportRendering.Image) customReportItem.GeneratedReportItem);

            // get the graphics output from R as image
            var rCode = (string) GetCustomPropertyValue(customReportItem.CustomProperties, "rGraphicsDevice:Code", "");
            var dpi =
                Convert.ToInt32(GetCustomPropertyValue(customReportItem.CustomProperties, "rGraphicsDevice:Dpi", 150));
            var renderTextOutputInstead =
                Convert.ToBoolean(GetCustomPropertyValue(customReportItem.CustomProperties,
                    "rGraphicsDevice:RenderTextOutputInstead", false));
            var fontFamily = customReportItem.Style.FontFamily.Value;
            var fontSizeEm = (int) (customReportItem.Style.FontSize.Value.ToInches()*dpi);
            fontSizeEm = fontSizeEm != 0 ? fontSizeEm : 16;
            var textOutputFont = new Font(fontFamily, fontSizeEm);
            var textOutputBrush = new SolidBrush(customReportItem.Style.Color.Value.ToColor());

            /* TODO: complete - renderTextOutputInstead setting. the font size does not work properly yet.
             * 
             * var rGraphicsDevice = new RGraphicsDevice(onDemandReportRenderingImage.Width.ToMillimeters(),
                onDemandReportRenderingImage.Height.ToMillimeters(), dpi, renderTextOutputInstead, textOutputFont,
                textOutputBrush);
             * */
            var rGraphicsDevice = new RGraphicsDevice(onDemandReportRenderingImage.Width.ToMillimeters(),
                onDemandReportRenderingImage.Height.ToMillimeters(), dpi, false, textOutputFont,
                textOutputBrush);
            string rScriptConsoleOutput;
            var rGraphicsOutput = rGraphicsDevice.GetDeviceOutputAsImage(rCode, out rScriptConsoleOutput);

            // update the customReportItem to use the device output image
            onDemandReportRenderingImage.ImageInstance.ImageData = ConvertImageToPngByteArray(rGraphicsOutput);
        }

        /// <summary>
        ///     Converts the specified image to a byte array which represents the image as PNG file.
        /// </summary>
        private static byte[] ConvertImageToPngByteArray(Image image)
        {
            // setup a memory stream
            var memoryStream = new MemoryStream();

            // save the image to the memory stream
            image.Save(memoryStream, ImageFormat.Png);

            // return the stream as an array
            return memoryStream.ToArray();
        }

        /// <summary>
        ///     Gets the value of the specified custom property.
        /// </summary>
        private static object GetCustomPropertyValue(CustomPropertyCollection customProperties, string name,
            object defaultValue)
        {
            // return the default value if we don't have the custom property
            if (customProperties == null || customProperties.Count == 0 || customProperties[name] == null)
            {
                return defaultValue;
            }

            // if we reach here, the custom property exists

            // get and return its value (done differently, depending whether it is an expression or not)
            var customProperty = customProperties[name];
            return customProperty.Value.IsExpression ? customProperty.Instance.Value : customProperty.Value.Value;
        }
    }
}