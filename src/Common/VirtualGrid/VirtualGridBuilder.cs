using yuseok.kim.dw2docs.Common.DwObjects;
using yuseok.kim.dw2docs.Common.Enums;
using yuseok.kim.dw2docs.Common.Exceptions;
using yuseok.kim.dw2docs.Common.Extensions;
using System.Drawing;

namespace yuseok.kim.dw2docs.Common.VirtualGrid
{
    public class VirtualGridBuilder
    {
        public int XThreshold { get; set; } = 3;
        public int YThreshold { get; set; } = 3;

        public VirtualGrid? Build(
            DwControlMatrix controlMatrix,
            byte dwProcessing,
            bool simplifyBands,
            bool discardFloating,
            out string? error)
        {
            /// TODO: take into account the band's height when adding padding to the rows
            error = null;
            try
            {
                var dwType = (DwType)dwProcessing;
                //VirtualGrid grid = new(DefineRows(controlMatrix), DefineColumns(controlMatrix));

                var cellRepo = controlMatrix.Controls.ToVirtualCellRepository();


                /// Make sure objects are not smaller than the threshold value
                foreach (var cell in cellRepo.Cells)
                {
                    if (cell.Width < XThreshold && !cell.Object.Floating)
                        cell.Width = XThreshold;
                }

                //var firstPass =
                //    RowFirstPass(cellRepo.CellsByY, controlMatrix.Bands);

                var (virtualRows,
                    rowsPerBand,
                    floatingControls) =
                    DefineRows(cellRepo.CellsByY, controlMatrix.Bands);



                foreach (var row in virtualRows)
                {
                    int i = 0;
                    var controlsToRemoveFromRow = new HashSet<VirtualCell>();

                    /// look for objects in the same row with overlapping X, and convert one of them to 
                    /// a floating control
                    foreach (var control in row.Objects)
                    {
                        for (int j = i; j < row.Objects.Count; ++j)
                        {
                            if (row.Objects[j] != control
                                && (
                                    (row.Objects[j].X >= control.X && row.Objects[j].X < control.RightBound)
                                    || (row.Objects[j].RightBound > control.X && row.Objects[j].RightBound <= control.RightBound)
                                ))
                            {
                                floatingControls.Add(row.Objects[j]);
                                controlsToRemoveFromRow.Add(row.Objects[j]);
                            }
                        }
                        i++;
                    }

                    foreach (var control in controlsToRemoveFromRow)
                    {
                        control.OwningRow = null;
                        control.OwningColumn = null;
                        row.Objects.Remove(control);
                    }
                }

                var virtualColumns = DefineColumns(cellRepo.CellsByX, floatingControls);

                switch (dwType)
                {
                    case DwType.Grid:
                        {
                            virtualColumns = RemoveFillerColumns(virtualColumns);
                            break;
                        }
                }

                if (discardFloating)
                {
                    floatingControls = floatingControls
                        .Where(x => x.Object.Floating)
                        .ToHashSet();
                }

                // validate integrity (rows and columns must contain the exact same controls)
                CheckIntegrity(virtualRows, virtualColumns);

                /// Add padding entities to allow for placing floating controls that overlap existing entities
                AddNecessaryPaddingColumns(ref virtualColumns, floatingControls);

                /// Make sure the band has the same height as it defined in the DW
                /// if not, insert rows to match
                int bottomMostBound;
                int missingPadding;
                RowDefinition? bottomMostRow = null;
                foreach (var band in controlMatrix.Bands)
                {
                    bottomMostBound = int.MinValue;
                    if (rowsPerBand.ContainsKey(band.Name))
                    {
                        foreach (var row in rowsPerBand[band.Name])
                        {
                            if (bottomMostBound < row.Bound)
                            {
                                bottomMostBound = row.Bound;
                                bottomMostRow = row;
                            }
                        }

                        if (bottomMostRow is not null)
                        {

                            missingPadding = band.Bound - bottomMostBound;
                            if (missingPadding > 0)
                            {
                                var newRow = new RowDefinition
                                {
                                    BandName = band.Name,
                                    IsFiller = true,
                                    IsPadding = true,
                                    Size = missingPadding,
                                    PreviousEntity = bottomMostRow,
                                    NextEntity = bottomMostRow.NextEntity
                                };
                                if (bottomMostRow.NextEntity is not null)
                                {
                                    bottomMostRow.NextEntity.PreviousEntity = newRow;
                                }

                                bottomMostRow.NextEntity = newRow;

                                virtualRows.Insert(virtualRows.IndexOf(bottomMostRow) + 1, newRow);
                                rowsPerBand[band.Name].Add(newRow);
                            }

                        }
                    }
                }



                var paddingRow = AddNecessaryPaddingRows(ref virtualRows, floatingControls);
                if (paddingRow is not null)
                {
                    var lastRow = paddingRow.PreviousEntity;
                    foreach (var (band, rows) in rowsPerBand)
                    {
                        if (lastRow != null && rows.Contains(lastRow))
                        {
                            rowsPerBand[band].Add(paddingRow);
                            break;
                        }
                    }
                }

                /// dissolve filler rows
                /// 

                var fillerRows = virtualRows.Where(row => row.IsFiller
                                                                && row.PreviousEntity is not null
                                                                && row.NextEntity is not null)
                    .ToList();

                foreach (var row in fillerRows)
                {

                    bool dissolved = false;
                    if (row.Objects.Count > 0 || row.FloatingObjects.Count > 0)
                    {
                        throw new Exception("Filler row has content");
                    }

                    if (row.PreviousEntity is not null)
                    {

                        // Compensate floating controls' height with the height of the row we're about to remove
                        foreach (var floatingControl in floatingControls)
                        {
                            if (floatingControl.LowerBound >= row.PreviousEntity.Offset && floatingControl.LowerBound <= row.PreviousEntity.Bound)
                            {
                                floatingControl.Height += row.Size;
                            }
                        }

                        row.PreviousEntity.Size += row.Size;
                        row.PreviousEntity.NextEntity = row.NextEntity;



                        dissolved = true;
                    }

                    if (row.NextEntity is not null)
                    {
                        row.NextEntity.PreviousEntity = row.PreviousEntity;
                        if (!dissolved)
                        {
                            // Compensate floating controls' height and Y with the height of the row we're about to remove
                            foreach (var floatingControl in floatingControls)
                            {
                                if (floatingControl.Y >= row.NextEntity.Offset && floatingControl.Y <= row.NextEntity.Bound)
                                {
                                    floatingControl.Height += row.Size;
                                    floatingControl.Y -= row.Size;
                                }
                            }

                            row.NextEntity.Size += row.Size;
                        }
                    }

                    if (row.BandName != null && rowsPerBand.ContainsKey(row.BandName))
                    {
                        rowsPerBand[row.BandName].Remove(row);
                    }
                    virtualRows.Remove(row);
                }

                // assign floating objects to their rows

                FloatingCellOffset? cellOffset;
                foreach (var floatingControl in floatingControls)
                {
                    cellOffset = MapFloatingControl(virtualColumns,
                        virtualRows,
                        floatingControl);
                    if (cellOffset is null)
                    {
                        throw new InvalidOperationException("Could not find an appropriate reference column for cell");
                    }


                    //floatingControl
                    //    .OwningRow!
                    cellOffset
                        .StartRow
                        .FloatingObjects
                        .Add(FloatingVirtualCell.FromVirtualCell(
                                floatingControl,
                                cellOffset));
                }


                VirtualGrid grid = new(virtualRows,
                    virtualColumns,
                    controlMatrix.Bands.DwBandsToBandRows(rowsPerBand),
                    cellRepo,
                    dwType);


                return grid;
            }
            catch (Exception e)
            {
                error = e.Message;
                throw;
            }
        }

        private static (IList<RowCandidate> fixedRows, IList<RowCandidate> floatingRows) GetRowCandidates(IList<RowCandidate> rows)
        {
            var ordered = rows
                .OrderByDescending(row => row.Objects.Count)
                .ThenBy(row => row.Height);

            var rowProcessedArray = new bool[rows.Count];

            var proposedRows = new List<RowCandidate>();
            var floatingRows = new List<RowCandidate>();

            int i = -1;
            foreach (var row in ordered)
            {
                ++i;
                if (rowProcessedArray[i])
                    continue;

                int j = -1;
                foreach (var rowToCompare in ordered)
                {
                    ++j;
                    if (row == rowToCompare || rowProcessedArray[j])
                        continue;

                    /// Rows overlap
                    /// TODO: introduce threshold 
                    if (row.Offset < rowToCompare.Bound && rowToCompare.Offset < row.Bound)
                    {
                        floatingRows.Add(rowToCompare);
                        rowProcessedArray[j] = true;
                    }

                }

                proposedRows.Add(row);
                rowProcessedArray[i] = true;
            }

            /// There should not be any more overlapping

            return (proposedRows, floatingRows);
        }

        private static void CheckIntegrity(
            ICollection<RowDefinition> rows,
            ICollection<ColumnDefinition> columns)
        {
            var map = new Dictionary<VirtualCell, byte>();

            foreach (var row in rows)
            {
                // Not taking into account floating objects
                foreach (var control in row.Objects)
                {
                    if (!map.ContainsKey(control))
                        map[control] = 0;
                    map[control]++;
                }
            }

            // All controls in the columns must exist as solid cells in the rows, 
            foreach (var col in columns)
            {
                foreach (var control in col.Objects)
                {
                    if (!map.ContainsKey(control))
                        map[control] = 0;
                    map[control]++;
                }
            }

            foreach (var (control, occurrences) in map)
            {
                if (occurrences != 2)
                {
                    throw new GridObjectCountInconsistencyException(control.Object);
                }
            }
        }

        private static IList<ColumnDefinition> RemoveFillerColumns(IList<ColumnDefinition> columns)
        {
            var newColumnList = new List<ColumnDefinition>();

            foreach (var column in columns)
            {
                if (column.IsFiller && column.Objects.Count == 0)
                {
                    column.RemoveFromChain();
                }
                else
                {
                    column.CalculateOffset();
                    newColumnList.Add(column);
                }
            }

            return newColumnList;

        }

        private static ColumnDefinition? AddNecessaryPaddingColumns(ref IList<ColumnDefinition> columns, ISet<VirtualCell> cells)
        {
            return AddNecessaryPaddingEntities(ref columns, cells, (cell) => cell.RightBound);
        }

        private static RowDefinition? AddNecessaryPaddingRows(ref IList<RowDefinition> rows, ISet<VirtualCell> cells)
        {
            var row = AddNecessaryPaddingEntities(ref rows, cells, (cell) => cell.LowerBound);
            if (row is not null)
                row.BandName = rows.Last().BandName;
            return row;
        }

        private static T? AddNecessaryPaddingEntities<T>(ref IList<T> entities,
                ISet<VirtualCell> cells,
                Func<VirtualCell, int> boundAccessor)
            where T : EntityDefinition<T>, new()
        {
            int furthestReachingControlBound = int.MinValue;
            int cellBound;
            foreach (var cell in cells)
            {
                cellBound = boundAccessor(cell);
                furthestReachingControlBound = cellBound > furthestReachingControlBound ? cellBound : furthestReachingControlBound;
            }

            var lastColumn = entities.Last();

            if (furthestReachingControlBound > lastColumn.Bound)
            {
                var newEntity = new T
                {
                    PreviousEntity = lastColumn,
                    Size = furthestReachingControlBound - lastColumn.Bound,
                };

                lastColumn.NextEntity = newEntity;
                entities.Add(newEntity);
                return newEntity;
            }

            return null;
        }


        /// <summary>
        /// Define the rows 
        /// </summary>
        /// <param name="controls">List of controls ordered by X. Items might be removed from this list if they overlap with other
        /// controls on Y</param>
        /// <param name="floatingControls">List of floating controls that will be excluded from calculations. Items might
        /// be added to this list if they overlap with other controls on Y</param>
        /// <returns></returns>
        private (IList<RowDefinition> rows,
            IDictionary<string, IList<RowDefinition>> rowsPerBand,
            ISet<VirtualCell> floatingObjects)
            DefineRows(
            SortedList<int, IList<VirtualCell>> objectsByY,
            IList<DwBand> bands)
        {
            var candidateRowMap = new Dictionary<int, IList<RowCandidate>>();
            Dictionary<string, HashSet<RowCandidate>> rowsPerBandDictionary = new();
            IDictionary<string, IList<RowDefinition>> rowdefinitionsPerBand = new Dictionary<string, IList<RowDefinition>>();

            foreach (var band in bands)
            {
                rowsPerBandDictionary[band.Name] = new();
                rowdefinitionsPerBand[band.Name] = new List<RowDefinition>();
            }

            /// Normalize y positions based on the YThreshold value
            var normalizedObjectsByY = new SortedList<int, IList<VirtualCell>>();
            int delta;
            foreach (var (y, objects) in objectsByY)
            {
                int previousY = normalizedObjectsByY.LastOrDefault().Key;
                delta = y - (previousY);

                if (delta <= YThreshold)
                {
                    if (!normalizedObjectsByY.ContainsKey(previousY))
                        normalizedObjectsByY[previousY] = new List<VirtualCell>();

                    foreach (var @object in objects)
                    {
                        @object.Y -= delta;
                    }
                    ((List<VirtualCell>)normalizedObjectsByY[previousY]).AddRange(objects);
                }
                else
                {
                    if (!normalizedObjectsByY.ContainsKey(y))
                        normalizedObjectsByY[y] = objects;
                    else
                        ((List<VirtualCell>)normalizedObjectsByY[y]).AddRange(objects);
                }

            }


            var floatingObjects = new HashSet<VirtualCell>();

            /// Create candidate rows from the objects. Objects starting in the
            /// same Y and having the same size will be aggregated.
            foreach (var (y, objects) in normalizedObjectsByY)
            {
                if (!candidateRowMap.ContainsKey(y))
                {
                    candidateRowMap[y] = new List<RowCandidate>();
                }

                foreach (var @object in objects)
                {
                    if (@object.Object.Floating || @object.Object.Band == "foreground" || @object.Object.Band == "background")
                    {
                        floatingObjects.Add(@object);
                        continue;
                    }
                    var candidateRow = candidateRowMap[y]
                        .Where(candidate => candidate.Offset == y)
                        .OrderBy(candidate => Math.Abs(candidate.Height - @object.Height))
                        .FirstOrDefault();

                    if (candidateRow is null || Math.Abs(candidateRow.Height - @object.Height) > YThreshold)
                    {
                        candidateRowMap[y].Add(candidateRow = new RowCandidate
                        {
                            Offset = y,
                            Height = @object.Height,
                            Objects = new List<VirtualCell>()
                            {
                                @object
                            },
                            Band = @object.Object.Band
                        });

                        rowsPerBandDictionary[@object.Object.Band].Add(candidateRow);
                    }
                    else
                    {
                        delta = candidateRow.Height - @object.Height;
                        @object.Height += delta;
                        candidateRow.Objects.Add(@object);
                    }

                }
            }

            var (fixedRows, floatingRows) =
                GetRowCandidates(candidateRowMap
                    .Values
                    .Where(rowList => rowList.Count > 0)
                    .SelectMany(row => row)
                    .ToList());

            var fixedRowsByY = fixedRows.OrderBy(row => row.Offset).ToList();
            var filledRows = new List<RowCandidate>();


            floatingObjects.AddRange(floatingRows.SelectMany(row => row.Objects));

            ///
            /// Fill up the empty space between each row
            /// 
            delta = 0;
            RowCandidate? previousRow = null;

            foreach (var rowCandidate in fixedRowsByY)
            {
                while ((delta = rowCandidate.Offset - (previousRow?.Bound ?? 0)) != 0)
                {
                    if (delta <= YThreshold)
                    {
                        rowCandidate.Offset -= delta;
                        rowCandidate.Height += delta;
                    }
                    else
                    {
                        bool adjustedPreviousRow = false;
                        int bandDelta = 0;
                        if (previousRow is not null && previousRow.Band != rowCandidate.Band)
                        {
                            var band = bands.Where(b => b.Name == previousRow.Band).First();

                            if ((bandDelta = band.Bound - previousRow.Bound) >= YThreshold)
                            {
                                previousRow = new RowCandidate()
                                {
                                    IsFiller = true,
                                    Band = previousRow.Band,
                                    Offset = previousRow.Bound,
                                    Height = bandDelta,
                                };
                                filledRows.Add(previousRow);
                            }
                            else
                            {
                                previousRow.Height += bandDelta;
                            }
                            adjustedPreviousRow = true;

                        }
                        previousRow = new RowCandidate()
                        {
                            Offset = previousRow?.Bound ?? 0,
                            Height = delta - bandDelta,
                            IsFiller = true,
                            Band = adjustedPreviousRow || previousRow is null ? rowCandidate.Band : previousRow.Band,
                        };
                        filledRows.Add(previousRow);
                    }

                }

                filledRows.Add(rowCandidate);
                previousRow = rowCandidate;
            }

            /// compensate difference between last row and end of band
            var lastRow = filledRows.Last();
            var lastBandDiff = bands.Where(b => b.Name == lastRow.Band).Single().Bound - lastRow.Bound;

            if ((lastBandDiff <= YThreshold))
            {
                lastRow.Height += lastBandDiff;
            }
            else
            {
                filledRows.Add(new RowCandidate
                {
                    IsFiller = true,
                    Band = lastRow.Band,
                    Height = lastBandDiff,
                    Offset = lastRow?.Bound ?? 0,
                });
            }

            var rowDefinitions = new List<RowDefinition>();

            RowDefinition? rowDef = null;
            /// Row candidates are normalized, we can convert them direclty
            /// 
            foreach (var row in filledRows)
            {
                rowDefinitions.Add(rowDef = new RowDefinition()
                {
                    Objects = row.Objects,
                    BandName = row.Band,
                    IsFiller = row.IsFiller,
                    PreviousEntity = rowDef,
                    Size = row.Height,
                });

                foreach (var @object in row.Objects)
                {
                    @object.OwningRow = rowDef;
                }

                if (rowDef.PreviousEntity is not null)
                    rowDef.PreviousEntity.NextEntity = rowDef;
                if (rowDef.BandName is null)
                {
                    throw new InvalidOperationException("Row doesn't belong to any band");
                }
                rowdefinitionsPerBand[rowDef.BandName].Add(rowDef);


            }

            /// Sort Row's controls by X
            foreach (var row in rowDefinitions)
            {
                ((List<VirtualCell>)(row.Objects)).Sort((a, b) => a.X - b.X);
            }

            return (rowDefinitions,
                rowdefinitionsPerBand,
                floatingObjects);
        }



        /// <summary>
        /// Define the rows 
        /// </summary>
        /// <param name="controls">List of controls ordered by X. Items might be removed from this list if they overlap with other
        /// controls on Y</param>
        /// <param name="floatingControls">List of floating controls that will be excluded from calculations. Items might
        /// be added to this list if they overlap with other controls on Y</param>
        /// <returns></returns>
        /// 
        private IList<ColumnDefinition> DefineColumns(
            SortedList<int, IList<VirtualCell>> controls,
            ISet<VirtualCell> floatingControls)
        {
            var normalizedList = new SortedDictionary<int, IList<VirtualCell>>();

            int rightmostBound

                = Math.Max(
                controls.Values
                    .SelectMany(c => c)
                    .Max(c => c.RightBound),
                floatingControls
                    .DefaultIfEmpty()
                    .Max(c => c?.RightBound ?? 0));

            /// Normalize controls based on the XThreshold
            int lastX = 0;
            int delta;
            foreach (var (x, controlSet) in controls)
            {
                if ((delta = x - normalizedList.LastOrDefault().Key) <= XThreshold && delta != 0)
                {
                    if (!normalizedList.ContainsKey(lastX))
                    {
                        normalizedList[lastX] = new List<VirtualCell>();
                    }

                    foreach (var control in controlSet)
                    {
                        control.X -= delta;
                    }

                    ((List<VirtualCell>)normalizedList[lastX]).AddRange(controlSet);
                }
                else
                {
                    normalizedList[x] = controlSet;
                    lastX = x;
                }

                lastX = x;
            }

            var normalizedFiltered = new SortedDictionary<int, IList<VirtualCell>>();
            /// Remove floating controls from the control list, and order them by Width
            foreach (var x in normalizedList.Keys)
            {
                normalizedFiltered[x] = normalizedList[x]
                    .Where(c => !floatingControls.Any(c1 => c1 == c))
                    .OrderBy(c => c.Width)
                    .ToList();
            }

            List<int> boundaries = new() { 0 };
            /// Calculate all boundaries of all controls, 
            foreach (var (x, controlSet) in normalizedFiltered)
            {
                boundaries.Add(x);

                foreach (var control in controlSet)
                {
                    boundaries.Add(control.RightBound);
                }
            }
            boundaries.Sort();
            boundaries = boundaries.Distinct().ToList();

            var columnDefinitions = new List<ColumnDefinition>();

            /// Assign a column into each boundary interval
            Dictionary<int, ColumnDefinition> columnOffsetIndex = new();
            Dictionary<int, ColumnDefinition> columnBoundIndex = new();
            ColumnDefinition? previousColumn = null;
            int previousBoundary = -1;
            foreach (var boundary in boundaries)
            {
                if (previousBoundary == -1)
                {
                    previousBoundary = boundary;
                    continue;
                }

                columnDefinitions.Add(previousColumn = new ColumnDefinition
                {
                    PreviousEntity = previousColumn,
                    IsFiller = true,
                    Size = boundary - previousBoundary,
                });
                columnOffsetIndex[previousBoundary] = previousColumn;
                columnBoundIndex[boundary] = previousColumn;

                previousBoundary = boundary;
            }

            for (int i = columnDefinitions.Count - 1; i > 0; --i)
            {
                if (columnDefinitions[i].PreviousEntity is not null)
                {
                    columnDefinitions[i].PreviousEntity!.NextEntity = columnDefinitions[i];
                }
            }

            List<VirtualCell> unmappedCells = new();
            /// Remap all controls to the columns
            foreach (var controlList in normalizedFiltered.Values)
            {
                foreach (var control in controlList)
                {
                    var startColumn = columnOffsetIndex[control.X];
                    var endColumn = columnBoundIndex[control.RightBound];

                    if (startColumn is null)
                        throw new InvalidOperationException("Could not find Offset Column for control");
                    if (endColumn is null)
                        throw new InvalidOperationException("Could not find Bound Column for control");

                    control.ColumnSpan = 1 + endColumn.IndexOffset - startColumn.IndexOffset;
                    control.OwningColumn = startColumn;
                    startColumn.Objects.Add(control);
                    startColumn.IsFiller = false;
                }
            }

            if (unmappedCells.Count > 0)
            {
                throw new InvalidOperationException("Controls not mapped to any column.");
            }


            /// Disolve columns that are smaller than threshold value
            var columnsToDelete = new List<ColumnDefinition>();
            foreach (var column in columnDefinitions)
            {

                if (column.Size < XThreshold)
                {
                    columnsToDelete.Add(column);
                    /// Adjust ColumnSpan values for all previous rows's objects
                    ColumnDefinition? columnRef = column.PreviousEntity;
                    int spanRef = 2;
                    if (columnRef is not null)
                    {
                        columnRef.Size += column.Size;
                    }
                    while (columnRef is not null)
                    {
                        foreach (var @object in columnRef.Objects)
                        {
                            if (@object.ColumnSpan >= spanRef)
                                --@object.ColumnSpan;
                        }

                        ++spanRef;
                        columnRef = columnRef.PreviousEntity;
                    }

                    if (column.NextEntity is null && column.Objects.Any())
                        throw new InvalidOperationException("Attempting to dissolve non-empty column without right-side replacement");

                    foreach (var @object in column.Objects)
                    {
                        @object.OwningColumn = column.NextEntity;
                        column.NextEntity!.Objects.Add(@object);
                        if (@object.ColumnSpan == 1)
                            throw new InvalidOperationException("Attempting to disolve column with unmovable object");
                        else
                            --@object.ColumnSpan;
                        @object.X = column.NextEntity.Offset;
                        @object.Width -= column.Size;
                    }

                    if (column.PreviousEntity is not null)
                    {
                        column.PreviousEntity.NextEntity = column.NextEntity;
                    }
                    if (column.NextEntity is not null)
                        column.NextEntity.PreviousEntity = column.PreviousEntity;

                    column.PreviousEntity = null;
                    column.NextEntity = null;
                    column.Objects.Clear();
                }
            }
            foreach (var column in columnsToDelete)
            {
                columnDefinitions.Remove(column);
            }

            var lastColumn = columnDefinitions.Last();
            lastColumn.RecalculateChainOffsets();
            int paddingRequired = rightmostBound - lastColumn.Bound;

            if (paddingRequired <= XThreshold)
            {
                lastColumn.Size += paddingRequired;

                /// Adjust all previous object's widths to match the column adjustment
                int colRef = 1;
                ColumnDefinition? currentColumn = lastColumn;
                while (currentColumn is not null)
                {
                    foreach (var @object in currentColumn.Objects)
                    {
                        if (@object.ColumnSpan >= colRef)
                            @object.Width += paddingRequired;
                    }

                    currentColumn = currentColumn.PreviousEntity;
                    ++colRef;
                }
            }
            else
            {
                var newColumn = new ColumnDefinition
                {
                    PreviousEntity = lastColumn,
                    Size = paddingRequired,
                    IsPadding = true,
                };
                columnDefinitions.Add(newColumn);

                lastColumn.NextEntity = newColumn;
            }

            return columnDefinitions;
        }

        private static FloatingCellOffset? MapFloatingControl(IList<ColumnDefinition> columns, IList<RowDefinition> rows, VirtualCell cell)
        {
            int distanceToStartColumn = int.MaxValue;
            int currentOffset;
            ColumnDefinition? startColumn = null;
            ColumnDefinition? endColumn;

            /// get left-bounding column
            foreach (var column in columns)
            {
                currentOffset = cell.X - column.Offset;
                if (currentOffset < distanceToStartColumn && currentOffset >= 0)
                {
                    distanceToStartColumn = currentOffset;
                    startColumn = column;
                }
            }

            if (startColumn is null)
            {
                return null;
            }

            endColumn = startColumn;
            short colSpan = 1;
            int distanceToEndColumn = int.MinValue;

            while (endColumn is not null)
            {
                if ((endColumn.Bound - (cell.RightBound)) >= 0)
                {
                    distanceToEndColumn = endColumn.Bound - cell.RightBound;
                    break;
                }

                endColumn = endColumn.NextEntity;
                ++colSpan;
            }


            if (endColumn is null)
            {
                throw new InvalidOperationException("Could not find rightmost column buffer for floating controls." +
                    " This should not happen");
            }



            ///// Get Y offsets
            //if (cell.OwningRow is null)
            //{
            //    throw new InvalidOperationException("Cell must have an owning row");
            //}

            int distanceToStartRow = int.MaxValue;
            RowDefinition? startRow = null;
            int distanceToCurrentRow;
            foreach (var row in rows)
            {
                distanceToCurrentRow = cell.Object.Y - row.Offset;
                if (distanceToCurrentRow < distanceToStartRow && distanceToCurrentRow >= 0)
                {
                    distanceToStartRow = distanceToCurrentRow;
                    startRow = row;
                }
            }

            if (startRow is null)
                throw new InvalidOperationException("Could not find start row");

            int distanceToEndRow = int.MaxValue;
            RowDefinition? endRow = startRow;
            short rowSpan = 1;
            while (endRow is not null)
            {
                if ((endRow.Bound - cell.LowerBound) >= 0)
                {
                    distanceToEndRow = endRow.Bound - cell.LowerBound;
                    break;
                }

                ++rowSpan;
                endRow = endRow.NextEntity;
            }

            if (endRow is null)
            {
                throw new InvalidOperationException("Could not find rightmost column buffer for floating controls." +
                    " This should not happen");
            }

            return new FloatingCellOffset(startColumn, startRow)
            {
                StartOffset = new Point(distanceToStartColumn, distanceToStartRow),
                EndOffset = new Point(distanceToEndColumn, distanceToEndRow),
                ColSpan = colSpan,
                RowSpan = rowSpan,
            };
        }

    }
}
