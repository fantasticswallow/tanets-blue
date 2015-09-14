using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml.Linq;
using artfulplace.XmlBlinker;

namespace tanets_blue_clients
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            cli = new tanets_blue.core.OneNoteClient();
            HTMLGenerator.GetBinaryObject = (x1, x2) => cli.GetBinaryObject(x1, x2);
                                             
        }

        private tanets_blue.core.OneNoteClient cli;

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            HTMLGenerator.IsResizeFont = checkBox.IsChecked.Value;
            var l = cli.GetCurrentId();
            var px = cli.GetPageContent(l);
            var xe = HTMLGenerator.CreateFromXml(px);
            var id = HTMLGenerator.PageId;
            id = id.Remove(id.LastIndexOf('}')).Remove(0,id.LastIndexOf('{') + 1);

            var xmlPath = System.IO.Path.Combine(PathTextBox.Text, id + ".html");
            HTMLGenerator.WriteXml(xmlPath, xe);
        }
               
    }
}
