// RichTextBoxHelper.cs in GLSense.Addin.Core
// Group H (Balance Configurator) addition - verbatim port of GLSense\Converters\
// RichTextBoxHelper.cs (FinalWorkingCode). Attached DependencyProperty that lets
// GLBalanceConfigurator.xaml bind RichTextBox.Document to GLConfiguratorViewModel.
// ParameterDisplayText (a FlowDocument), which WPF's RichTextBox.Document property does
// not natively support binding to.
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;

namespace GLSense.Addin.Core.Converters
{
    public static class RichTextBoxHelper
    {
        public static readonly DependencyProperty BindableDocumentProperty =
            DependencyProperty.RegisterAttached(
                "BindableDocument",
                typeof(FlowDocument),
                typeof(RichTextBoxHelper),
                new PropertyMetadata(null, OnBindableDocumentChanged));

        public static void SetBindableDocument(DependencyObject obj, FlowDocument value)
            => obj.SetValue(BindableDocumentProperty, value);

        public static FlowDocument GetBindableDocument(DependencyObject obj)
            => (FlowDocument)obj.GetValue(BindableDocumentProperty);

        private static void OnBindableDocumentChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is RichTextBox rtb)
            {
                // Avoid re-assigning the same document instance which can cause
                // re-entrancy/layout loops when the FlowDocument updates and
                // the binding tries to set the property again. Only replace the
                // RichTextBox.Document when the instance actually changes.
                var newDoc = e.NewValue as FlowDocument ?? new FlowDocument();
                if (!ReferenceEquals(rtb.Document, newDoc))
                {
                    // Assign new document instance
                    rtb.Document = newDoc;
                }
            }
        }
    }
}
