using System;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Projection;

namespace BJVSExtension.Utilities
{
    internal enum CodeRegionType
    {
        None,
        /// <summary>
        ///  { }
        /// </summary>
        Block,
        /// <summary>
        /// 注释
        /// </summary>
        Comment,
        /// <summary>
        /// #region #if #else #endif
        /// </summary>
        ProProcessor,
        /// <summary>
        /// using
        /// </summary>
        Using,
        /// <summary>
        ///  case default
        /// </summary>
        Switch,
        /// <summary>
        /// 函数
        /// </summary>
        Method,
    }

    class CodeRegin
    {
        public CodeRegin(SnapshotPoint start, CodeRegionType regionType, ITextEditorFactoryService editorFactory, IProjectionBufferFactoryService bufferFactory)
        {
            StartPoint = start;
            EndPoint = start;
            RegionType = regionType;
            EditorFactory = editorFactory;
            BufferFactory = bufferFactory;
        }

#if DEBUG

        public string DebugStartLineText => StartLine.GetText();

        public string DebugEndLineText => Complete ? EndLine.GetText() : "";

        public string DebugRegionText => Complete ? StartPoint.Snapshot.GetText(StartPoint.Position, EndPoint.Position - StartPoint.Position) : "";

#endif


        ITextEditorFactoryService EditorFactory;
        IProjectionBufferFactoryService BufferFactory;


        public CodeRegionType RegionType = CodeRegionType.None;

        public bool Complete = false;

        public SnapshotPoint StartPoint;

        public SnapshotPoint EndPoint;

        public bool StartsFromLastLine = false;

        public ITextSnapshotLine StartLine => StartPoint.GetContainingLine();

        public ITextSnapshotLine EndLine => EndPoint.GetContainingLine();

        public int SpanIndex = -1;

        public string StartSpanText = null;

        public SnapshotSpan ToSnapshotSpan()
        {
            return new SnapshotSpan(this.StartPoint, this.EndPoint);
        }

        public override string ToString()
        {
            return $"{RegionType},{(Complete ? "Complete" : "Open")}, {StartPoint.Position}-{EndPoint.Position} Line:{StartLine.LineNumber}-{EndLine.LineNumber}";
        }
    }
}
