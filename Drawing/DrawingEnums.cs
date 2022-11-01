using System;

namespace SpreadsheetLight.Drawing
{
    /// <summary>
    /// Specifies how the end points are joined.
    /// </summary>
    public enum SLLineJoinValues
    {
        /// <summary>
        /// For rounded joins.
        /// </summary>
        Round = 1,
        /// <summary>
        /// For bevelled joins.
        /// </summary>
        Bevel,
        /// <summary>
        ///  For miter joins (square edges).
        /// </summary>
        Miter
    }

    /// <summary>
    /// Specifies the size.
    /// </summary>
    public enum SLLineSizeValues
    {
        /// <summary>
        /// Size 1. Width is small, length is small.
        /// </summary>
        Size1 = 1,
        /// <summary>
        /// Size 2. Width is small, length is medium.
        /// </summary>
        Size2,
        /// <summary>
        /// Size 3. Width is small, length is large.
        /// </summary>
        Size3,
        /// <summary>
        /// Size 4. Width is medium, length is small.
        /// </summary>
        Size4,
        /// <summary>
        /// Size 5. Width is medium, length is medium.
        /// </summary>
        Size5,
        /// <summary>
        /// Size 6. Width is medium, length is large.
        /// </summary>
        Size6,
        /// <summary>
        /// Size 7. Width is large, length is small.
        /// </summary>
        Size7,
        /// <summary>
        /// Size 8. Width is large, length is medium.
        /// </summary>
        Size8,
        /// <summary>
        /// Size 9. Width is large, length is large.
        /// </summary>
        Size9
    }

    /// <summary>
    /// Built-in gradient preset colors.
    /// </summary>
    public enum SLGradientPresetValues
    {
        /// <summary>
        /// Early Sunset
        /// </summary>
        EarlySunset = 0,
        /// <summary>
        /// Late Sunset
        /// </summary>
        LateSunset,
        /// <summary>
        /// Nightfall
        /// </summary>
        Nightfall,
        /// <summary>
        /// Daybreak
        /// </summary>
        Daybreak,
        /// <summary>
        /// Horizon
        /// </summary>
        Horizon,
        /// <summary>
        /// Desert
        /// </summary>
        Desert,
        /// <summary>
        /// Ocean
        /// </summary>
        Ocean,
        /// <summary>
        /// Calm Water
        /// </summary>
        CalmWater,
        /// <summary>
        /// Fire
        /// </summary>
        Fire,
        /// <summary>
        /// Fog
        /// </summary>
        Fog,
        /// <summary>
        /// Moss
        /// </summary>
        Moss,
        /// <summary>
        /// Peacock
        /// </summary>
        Peacock,
        /// <summary>
        /// Wheat
        /// </summary>
        Wheat,
        /// <summary>
        /// Parchment
        /// </summary>
        Parchment,
        /// <summary>
        /// Mahogany
        /// </summary>
        Mahogany,
        /// <summary>
        /// Rainbow
        /// </summary>
        Rainbow,
        /// <summary>
        /// Rainbow II
        /// </summary>
        Rainbow2,
        /// <summary>
        /// Gold
        /// </summary>
        Gold,
        /// <summary>
        /// Gold II
        /// </summary>
        Gold2,
        /// <summary>
        /// Brass
        /// </summary>
        Brass,
        /// <summary>
        /// Chrome
        /// </summary>
        Chrome,
        /// <summary>
        /// Chrome II
        /// </summary>
        Chrome2,
        /// <summary>
        /// Silver
        /// </summary>
        Silver,
        /// <summary>
        /// Sapphire
        /// </summary>
        Sapphire
    }

    /// <summary>
    /// Specifies the direction for a radial or rectangular gradient type.
    /// </summary>
    public enum SLGradientDirectionValues
    {
        /// <summary>
        /// From center to top left corner.
        /// </summary>
        CenterToTopLeftCorner = 0,
        /// <summary>
        /// From center to top right corner.
        /// </summary>
        CenterToTopRightCorner,
        /// <summary>
        /// From center outwards.
        /// </summary>
        Center,
        /// <summary>
        /// From center to bottom left corner.
        /// </summary>
        CenterToBottomLeftCorner,
        /// <summary>
        /// From center to bottom right corner.
        /// </summary>
        CenterToBottomRightCorner
    }

    /// <summary>
    /// Built-in shadow preset values.
    /// </summary>
    public enum SLShadowPresetValues
    {
        /// <summary>
        /// None
        /// </summary>
        None = 0,
        /// <summary>
        /// Outer Diagonal Bottom Right
        /// </summary>
        OuterDiagonalBottomRight,
        /// <summary>
        /// Outer Bottom
        /// </summary>
        OuterBottom,
        /// <summary>
        /// Outer Diagonal Bottom Left
        /// </summary>
        OuterDiagonalBottomLeft,
        /// <summary>
        /// Outer Right
        /// </summary>
        OuterRight,
        /// <summary>
        /// Outer Center
        /// </summary>
        OuterCenter,
        /// <summary>
        /// Outer Left
        /// </summary>
        OuterLeft,
        /// <summary>
        /// Outer Diagonal Top Right
        /// </summary>
        OuterDiagonalTopRight,
        /// <summary>
        /// Outer Top
        /// </summary>
        OuterTop,
        /// <summary>
        /// Outer Diagonal Top Left
        /// </summary>
        OuterDiagonalTopLeft,
        /// <summary>
        /// Inner Diagonal Top Left
        /// </summary>
        InnerDiagonalTopLeft,
        /// <summary>
        /// Inner Top
        /// </summary>
        InnerTop,
        /// <summary>
        /// Inner Diagonal Top Right
        /// </summary>
        InnerDiagonalTopRight,
        /// <summary>
        /// Inner Left
        /// </summary>
        InnerLeft,
        /// <summary>
        /// Inner Center
        /// </summary>
        InnerCenter,
        /// <summary>
        /// Inner Right
        /// </summary>
        InnerRight,
        /// <summary>
        /// Inner Diagonal Bottom Left
        /// </summary>
        InnerDiagonalBottomLeft,
        /// <summary>
        /// Inner Bottom
        /// </summary>
        InnerBottom,
        /// <summary>
        /// Inner Diagonal Bottom Right
        /// </summary>
        InnerDiagonalBottomRight,
        /// <summary>
        /// Perspective Diagonal Upper Left
        /// </summary>
        PerspectiveDiagonalUpperLeft,
        /// <summary>
        /// Perspective Diagonal Upper Right
        /// </summary>
        PerspectiveDiagonalUpperRight,
        /// <summary>
        /// Perspective Below
        /// </summary>
        PerspectiveBelow,
        /// <summary>
        /// Perspective Diagonal Lower Left
        /// </summary>
        PerspectiveDiagonalLowerLeft,
        /// <summary>
        /// Perspective Diagonal Lower Right
        /// </summary>
        PerspectiveDiagonalLowerRight
    }

    /// <summary>
    /// Vertical text alignment.
    /// </summary>
    public enum SLTextVerticalAlignment
    {
        /// <summary>
        /// Top
        /// </summary>
        Top = 0,
        /// <summary>
        /// Middle
        /// </summary>
        Middle,
        /// <summary>
        /// Bottom
        /// </summary>
        Bottom,
        /// <summary>
        /// Top Centered
        /// </summary>
        TopCentered,
        /// <summary>
        /// Middle Centered
        /// </summary>
        MiddleCentered,
        /// <summary>
        /// Bottom Centered
        /// </summary>
        BottomCentered
    }

    /// <summary>
    /// Horizontal text alignment.
    /// </summary>
    public enum SLTextHorizontalAlignment
    {
        /// <summary>
        /// Right
        /// </summary>
        Right = 0,
        /// <summary>
        /// Center
        /// </summary>
        Center,
        /// <summary>
        /// Left
        /// </summary>
        Left,
        /// <summary>
        /// Right Middle
        /// </summary>
        RightMiddle,
        /// <summary>
        /// Center Middle
        /// </summary>
        CenterMiddle,
        /// <summary>
        /// Left Middle
        /// </summary>
        LeftMiddle
    }
}
