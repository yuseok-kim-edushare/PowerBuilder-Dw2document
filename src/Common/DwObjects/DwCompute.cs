using yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes;

namespace yuseok.kim.dw2docs.Common.DwObjects
{
    /// <summary>
    /// Represents a PowerBuilder DataWindow compute object
    /// </summary>
    public class DwCompute : IDwObject
    {
        /// <summary>
        /// Gets the name of the PowerBuilder DataWindow compute object
        /// </summary>
        public string Name { get; }
        
        /// <summary>
        /// Gets a value indicating whether this compute object is floating in the datawindow
        /// </summary>
        public bool Floating { get; }

        /// <summary>
        /// Gets the attributes of the compute object
        /// </summary>
        public DwComputeAttributes Attributes { get; } = new DwComputeAttributes();

        /// <summary>
        /// Initializes a new instance of the <see cref="DwCompute"/> class
        /// </summary>
        /// <param name="name">The name of the compute object</param>
        /// <param name="floating">Whether the compute object is floating</param>
        public DwCompute(string name, bool floating = false)
        {
            Name = name;
            Floating = floating;
        }
    }
}