using System.ComponentModel;
using System.ComponentModel.Design;
using System.Drawing;
using System.Linq;
using Microsoft.ReportDesigner;
using Microsoft.ReportingServices.Interfaces;
using Microsoft.ReportingServices.RdlObjectModel;

namespace SSRS.CustomReportItems.Designer
{
    /// <summary>
    ///     Designer class for the RGraphicsDevice Custom Report Item.
    /// </summary>
    [CustomReportItem("RGraphicsDevice")]
    [LocalizedName("R Graphics Device")]
    [Description("Shows the graphics device output of an arbitrary R code in Reporting Services.")]
    [ToolboxBitmap(typeof (RGraphicsDeviceDesigner), "RGraphicsDeviceDesigner.ico")]
    public class RGraphicsDeviceDesigner : CustomReportItemDesigner
    {
        /// <summary>
        ///     Private backing field.
        /// </summary>
        private IComponentChangeService _changeService;

        /// <summary>
        ///     Sets the default size.
        /// </summary>
        public override ItemSize DefaultSize
        {
            get { return new ItemSize(new ReportSize("90mm"), new ReportSize("80mm")); }
        }

        /// <summary>
        ///     Specifies the R code that produces the output on the graphics device.
        /// </summary>
        [Browsable(true)]
        [Category("R")]
        [Description("Specifies the R code that produces the output on the graphics device.")]
        public ReportExpression Code
        {
            get { return GetCustomProperty("rGraphicsDevice:Code"); }
            set
            {
                SetCustomProperty("rGraphicsDevice:Code", (string) value);
                RaiseComponentChanged();
                Invalidate();
            }
        }

        /// <summary>
        ///     Specifies the resolution of the graphics device.
        /// </summary>
        [Browsable(true)]
        [Category("R")]
        [Description("Specifies the resolution of the graphics device (in dpi).")]
        [TypeConverter(typeof (ReportExpressionConverter<int>))]
        [ReportExpressionDefaultValue(typeof (int), 150)]
        public ReportExpression<int> Dpi
        {
            get
            {
                var dpiAsString = GetCustomProperty("rGraphicsDevice:Dpi");
                dpiAsString = string.IsNullOrWhiteSpace(dpiAsString) || dpiAsString == "0" ? "150" : dpiAsString;
                return new ReportExpression<int>(dpiAsString);
            }
            set
            {
                SetCustomProperty("rGraphicsDevice:Dpi", (string) value);
                RaiseComponentChanged();
                Invalidate();
            }
        }

        /* TODO: remove once the RenderTextOutputInstead feature works properly
        /// <summary>
        ///     Specifies if the text output of R should be rendered instead of the graphics device output.
        /// </summary>
        [Browsable(true)]
        [Category("R")]
        [Description("Specifies if the text output of R should be rendered instead of the graphics device output.")]
        [TypeConverter(typeof (ReportExpressionConverter<bool>))]
        [ReportExpressionDefaultValue(typeof (bool), false)]
        public ReportExpression<bool> RenderTextOutputInstead
        {
            get
            {
                var renderTextOutputInsteadAsString = GetCustomProperty("rGraphicsDevice:RenderTextOutputInstead");
                var renderTextOutputInstead = !string.IsNullOrWhiteSpace(renderTextOutputInsteadAsString) &&
                                              Convert.ToBoolean(renderTextOutputInsteadAsString);
                return new ReportExpression<bool>(renderTextOutputInstead);
            }
            set
            {
                SetCustomProperty("rGraphicsDevice:RenderTextOutputInstead", (string) value);
                RaiseComponentChanged();
                Invalidate();
            }
        }
        */

        /// <summary>
        ///     Gets or sets the <see cref="IComponentChangeService" /> instance required by Visual Studio.
        /// </summary>
        /// <returns></returns>
        public IComponentChangeService ChangeService
        {
            get
            {
                return _changeService ??
                       (_changeService = (IComponentChangeService) Site.GetService(typeof (IComponentChangeService)));
            }
        }

        /// <summary>
        ///     Notifies the site's <see cref="IComponentChangeService" /> that the component has changed.
        /// </summary>
        private void RaiseComponentChanged()
        {
            ChangeService.OnComponentChanged(this, null, null, null);
        }

        /// <summary>
        ///     Gets the value of the specified custom property.
        /// </summary>
        private string GetCustomProperty(string propertyName)
        {
            return
                (from property in CustomProperties where property.Key == propertyName select property.Value)
                    .FirstOrDefault();
        }

        /// <summary>
        ///     Sets the value of the specified custom property.
        /// </summary>
        private void SetCustomProperty(string propertyName, string value)
        {
            if (!CustomProperties.ContainsKey(propertyName))
            {
                CustomProperties.Add(propertyName, value);
            }
            else
            {
                CustomProperties[propertyName] = value;
            }
        }

        /// <summary>
        ///     Draws the content of the designer component in Visual Studio.
        /// </summary>
        public override void Draw(Graphics graphics, ReportItemDrawParams reportItemDrawParameters)
        {
            // initialization
            var blackSolidBrush = new SolidBrush(Color.Black);
            var graySolidBrush = new SolidBrush(Color.Gray);
            const float margin = 3f;

            // draw the name of the report item
            var nameText = Name;
            var nameFont = new Font("Arial", 10);
            graphics.DrawString(nameText, nameFont, blackSolidBrush, new PointF(margin, margin));

            // draw the name of the report item
            var codeText = Code.ToString();
            var codeFont = new Font("Courier New", 8);
            var codeYPos = margin + graphics.MeasureString(nameText, nameFont).Height + 8f;
            graphics.DrawString(codeText, codeFont, graySolidBrush, new PointF(margin, codeYPos));
        }
    }
}