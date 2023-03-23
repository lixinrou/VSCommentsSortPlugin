﻿using System;
using Microsoft;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;

namespace BJVSExtension
{
    public static class TextViewHelper
    {
        public static IWpfTextView GetActiveTextView()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            IComponentModel componentModel = Package.GetGlobalService(typeof(SComponentModel)) as IComponentModel;
            if (componentModel == null)
                return null;
            
            IVsTextView activeView = GetActiveView();
            IVsEditorAdaptersFactoryService vsEditorAdapter = componentModel.GetService<IVsEditorAdaptersFactoryService>();
            return vsEditorAdapter.GetWpfTextView(activeView);
        }

        public static void ReplaceText(IWpfTextView textView, int startPosition, int length, string newText)
            => ReplaceText(textView, new Span(startPosition, length), newText);

        public static void ReplaceText(IWpfTextView textView, Span replaceSpan, string newText)
        {
            using (ITextEdit edit = textView.TextBuffer.CreateEdit())
            {
                edit.Replace(replaceSpan, newText);
                edit.Apply();
            }
        }

        /// <summary>
        /// 获取选中的起始结束行号 
        /// </summary>
        /// <param name="textView"></param>
        /// <returns></returns>
        public static (int startLineNo, int endLineNo) GetSelectedLineNumbers(IWpfTextView textView)
        {
            // Get positions of the first and last line in the selection
            int startSelectionPosition = textView.Selection.Start.Position.Position;
            int endSelectionPosition = textView.Selection.End.Position.Position;

            ITextSnapshot textSnapshot = textView.TextSnapshot;
            int startLineNumber = textSnapshot.GetLineNumberFromPosition(startSelectionPosition);
            int endLineNumber = textSnapshot.GetLineNumberFromPosition(endSelectionPosition);

            return (startLineNumber, endLineNumber);
        }

        /// <summary>
        /// 获取当前IVsTextView
        /// </summary>
        /// <returns></returns>
        private static IVsTextView GetActiveView()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            IVsTextManager vsTextManager = ServiceProvider.GlobalProvider.GetService(typeof(SVsTextManager)) as IVsTextManager;
            Assumes.Present(vsTextManager);

            ErrorHandler.ThrowOnFailure(vsTextManager.GetActiveView(1, null, out IVsTextView activeView));
            return activeView;
        }
        
    }
}
