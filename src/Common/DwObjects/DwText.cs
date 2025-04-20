using yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes;

namespace yuseok.kim.dw2docs.Common.DwObjects
{
    /// <summary>
    /// Represents a PowerBuilder DataWindow text object
    /// </summary>
    public class DwText : IDwObject
    {
        /// <summary>
        /// Gets the name of the PowerBuilder DataWindow text
        /// </summary>
        public string Name { get; }
        
        /// <summary>
        /// Gets a value indicating whether this text is floating in the datawindow
        /// </summary>
        public bool Floating { get; }

        /// <summary>
        /// Gets the attributes of the text
        /// </summary>
        public DwTextAttributes Attributes { get; } = new DwTextAttributes();

        /// <summary>
        /// Initializes a new instance of the <see cref="DwText"/> class
        /// </summary>
        /// <param name="name">The name of the text object</param>
        /// <param name="floating">Whether the text is floating</param>
        public DwText(string name, bool floating = false)
        {
            Name = name;
            Floating = floating;
        }
    }
} 