namespace PdfScriptTool
{
    internal enum AdvancedProgressBarDisplayText
    {
        Percent,
        Text
    }

    internal class AdvancedProgressBar : System.Windows.Forms.ProgressBar
    {
        public AdvancedProgressBarDisplayText DisplayStyle { get; set; }
        public override string Text { get; set; }

        public AdvancedProgressBar()
        {
            SetStyle(
                System.Windows.Forms.ControlStyles.UserPaint |
                System.Windows.Forms.ControlStyles.AllPaintingInWmPaint,
                true);
        }

        protected override void OnPaint(System.Windows.Forms.PaintEventArgs e)
        {
            var rectangle = ClientRectangle;
            var graphics = e.Graphics;
            System.Windows.Forms.ProgressBarRenderer.DrawHorizontalBar(
                graphics,
                rectangle);
            rectangle.Inflate(-3, -3);
            if (Value > 0)
            {
                var clip = new System.Drawing.Rectangle(
                    rectangle.X,
                    rectangle.Y,
                    (int)System.Math.Round(
                        ((float)Value / Maximum) * rectangle.Width),
                    rectangle.Height);
                if(System.Windows.Forms.Application.RenderWithVisualStyles)
                {
                    System.Windows.Forms.ProgressBarRenderer.DrawHorizontalChunks(
                    graphics,
                    clip);
                }
            }
            string text = DisplayStyle ==
                AdvancedProgressBarDisplayText.Percent ?
                Value.ToString() + '%' : Text;
            using (var font = new System.Drawing.Font(
                System.Drawing.FontFamily.GenericSansSerif,
                8))
            {
                var labelLength = graphics.MeasureString(text, font);
                var labelStartLocation = new System.Drawing.Point(
                    System.Convert.ToInt32(
                        (Width / 2) - labelLength.Width / 2),
                    System.Convert.ToInt32(
                        (Height / 2) - labelLength.Height / 2));
            }
        }
    }
}