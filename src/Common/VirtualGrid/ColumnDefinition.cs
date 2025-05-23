﻿using System.Text;

namespace yuseok.kim.dw2docs.Common.VirtualGrid
{
    public class ColumnDefinition : EntityDefinition<ColumnDefinition>
    {
        /// <summary>
        /// Splits the column into two at a point inside the column and remaps relationships
        /// </summary>
        /// <param name="at">The point at which to split the cell, relative value starting from the left edge</param>
        /// <returns>The resulting new column created by the split</returns>
        public ColumnDefinition Split(int at)
        {
            if (at > Size)
            {
                throw new ArgumentException("Specified split point is outside the column");
            }

            var originalSize = Size;

            Size = at;
            var newColumn = new ColumnDefinition()
            {
                Size = originalSize - at,
                PreviousEntity = this,
                NextEntity = NextEntity,
            };

            if (NextEntity is not null)
                NextEntity.PreviousEntity = newColumn;

            NextEntity = newColumn;

            // Update IndexOffset on columns ahead
            var refCol = NextEntity;
            int currentIndex = IndexOffset;
            while (refCol is not null)
            {
                refCol.IndexOffset = currentIndex++;
                refCol = refCol.NextEntity;
            }

            int reference = 1;

            refCol = this;
            while (refCol is not null)
            {
                foreach (var cell in refCol.Objects)
                {

                    if (cell.ColumnSpan >= reference)
                    {
                        ++cell.ColumnSpan;
                    }
                }

                ++reference;
                refCol = refCol.PreviousEntity;
            }

            return newColumn;
        }

        public override string? ToString()
        {
            var sb = new StringBuilder();

            sb.Append($"[X={Offset}]");
            sb.Append($"[Width={Size}] ");
            sb.Append($"[Controls={{{Objects.Aggregate(string.Empty, (str, obj) => $"{(str == string.Empty ? str : str + ", ")} {obj}")}}}]");
            sb.Append($"[Floating={{{FloatingObjects.Aggregate(string.Empty, (str, obj) => $"{(str == string.Empty ? str : str + ", ")} {obj}")}}}]");

            return sb.ToString();
        }
    }
}
