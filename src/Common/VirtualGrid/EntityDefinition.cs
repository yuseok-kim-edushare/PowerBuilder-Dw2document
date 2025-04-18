namespace yuseok.kim.dw2docs.Common.VirtualGrid
{
    public abstract class EntityDefinition<T>
        where T : EntityDefinition<T>
    {
        private T? _previousEntity;

        public T? PreviousEntity
        {
            get { return _previousEntity; }
            set
            {
                if (_previousEntity != value)
                {
                    _previousEntity = value;
                    CalculateOffset();
                }
            }
        }

        public T? NextEntity { get; set; }
        public int Offset { get; protected set; }
        /// <summary>
        /// 1 indexed value that indicates the ordinal position of the entity
        /// in the chain
        /// </summary>
        public int IndexOffset { get; protected set; }
        public int Size { get; set; }
        public int Bound => Offset + Size;

        public bool IsEmpty => Objects.Count == 0 && FloatingObjects.Count == 0;
        public IList<VirtualCell> Objects { get; set; }
        public IList<FloatingVirtualCell> FloatingObjects { get; set; }
        /// <summary>
        /// Whether this entity was inserted to fill empty space
        /// </summary>
        public bool IsFiller { get; set; }
        /// <summary>
        /// Whether this entity was inserted to add anchors for floating controls
        /// when such controls go over the existing entities
        /// </summary>
        public bool IsPadding { get; set; }

        public EntityDefinition()
        {
            Objects = new List<VirtualCell>();
            FloatingObjects = new List<FloatingVirtualCell>();
        }

        public virtual void RemoveCell(VirtualCell cell)
        {
            Objects.Remove(cell);
            FloatingObjects.Remove(FloatingObjects.Where(o => o == cell).Single());
        }

        public void CalculateOffset()
        {
            int offset = 0;
            int indexOffset = 0;
            T? previousEntity = PreviousEntity;

            while (previousEntity is not null)
            {
                offset += previousEntity.Size;
                previousEntity = previousEntity.PreviousEntity;
                ++indexOffset;
            }
            Offset = offset;
            IndexOffset = indexOffset;

        }

        public void RecalculateChainOffsets()
        {

            /// Get the first entity in the chain
            EntityDefinition<T>? currentEntity = this;
            while (currentEntity.PreviousEntity is not null)
            {
                currentEntity = currentEntity.PreviousEntity;
            }

            int i = 0;
            while (currentEntity is not null)
            {
                currentEntity.Offset = (currentEntity.PreviousEntity is null
                    ? 0
                    : currentEntity.PreviousEntity.Offset + currentEntity.PreviousEntity.Size);
                currentEntity.IndexOffset = currentEntity.PreviousEntity is null
                    ? 0
                    : ++i;
                currentEntity = currentEntity.NextEntity;
            }
        }

        public void RemoveFromChain()
        {
            if (Objects.Any())
            {
                throw new InvalidOperationException("Cannot remove an entity with controls in it");
            }
            if (PreviousEntity is not null)
            {
                PreviousEntity.NextEntity = NextEntity;
            }

            if (NextEntity is not null)
            {
                NextEntity.PreviousEntity = PreviousEntity;
            }

            PreviousEntity = null;
            NextEntity = null;
        }
    }
}
