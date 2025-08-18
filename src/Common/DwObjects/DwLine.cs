using yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes;

namespace yuseok.kim.dw2docs.Common.DwObjects
{
    /// <summary>
    /// Represents a PowerBuilder DataWindow line object
    /// </summary>
    public class DwLine : IDwObject
    {
        /// <summary>
        /// Gets the name of the PowerBuilder DataWindow line
        /// </summary>
        public string Name { get; }
        
        /// <summary>
        /// Gets a value indicating whether this line is floating in the datawindow
        /// </summary>
        public bool Floating { get; }

        /// <summary>
        /// Gets the attributes of the line
        /// </summary>
        public DwLineAttributes Attributes { get; } = new DwLineAttributes();

        /// <summary>
        /// Initializes a new instance of the <see cref="DwLine"/> class
        /// </summary>
        /// <param name="name">The name of the line object</param>
        /// <param name="floating">Whether the line is floating</param>
        public DwLine(string name, bool floating = false)
        {
            Name = name;
            Floating = floating;
        }
    }
}