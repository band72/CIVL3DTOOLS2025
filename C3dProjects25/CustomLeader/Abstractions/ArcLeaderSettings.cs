namespace RCS.CustomLeader.Abstractions
{
    public sealed class ArcLeaderSettings
    {
        public string HeadBlockName { get; set; } = "RCS_LEADER_HEAD";
        public double HeadScale { get; set; } = 1.0;
        public double HeadRotationOffsetDeg { get; set; } = 0.0;

        public string ArcLayer { get; set; } = "C-ANNO-LEAD";
        public string HeadLayer { get; set; } = "C-ANNO-LEAD";
        public string TextLayer { get; set; } = "C-ANNO-NOTE";

        public string TextStyleName { get; set; } = "Standard";
        public double TextHeight { get; set; } = 0.10;
        public double BoxOffset { get; set; } = 0.20;
        public double BoxPadding { get; set; } = 0.08;

        public bool UseBackgroundMask { get; set; } = true;
        public bool CreateRectangularFrame { get; set; } = true;
        public string DefaultText { get; set; } = "INTEX";
    }
}
