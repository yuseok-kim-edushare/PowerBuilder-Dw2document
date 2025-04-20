namespace yuseok.kim.dw2docs.Common.DwObjects
{
    /// <summary>
    /// Interface for PowerBuilder DataWindow objects
    /// </summary>
    public interface IDwObject
    {
        /// <summary>
        /// Gets the name of the PowerBuilder DataWindow object
        /// </summary>
        string Name { get; }
        
        /// <summary>
        /// Gets a value indicating whether this object is floating in the datawindow
        /// </summary>
        bool Floating { get; }
    }
} 