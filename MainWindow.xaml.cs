using System;
using System.Collections.Generic;
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

using System.Printing;
using System.IO.Packaging;

namespace WPFPrintPages
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // デフォルトプリンターで印刷します。
            var server = new LocalPrintServer();
            var queue = server.GetPrintQueues().FirstOrDefault(_ => _.Name == "Microsoft Print to PDF");

            // 用紙サイズ(A4)と印刷方向(縦)を設定します。
            var ticket = queue.DefaultPrintTicket;
            ticket.PageMediaSize = new PageMediaSize(PageMediaSizeName.ISOA4);
            ticket.PageOrientation = PageOrientation.Portrait;

            // 余白の上と左位置を取得します。
            var area = queue.GetPrintCapabilities().PageImageableArea;
            var originHeight = area.OriginHeight;
            var originWidth = area.OriginWidth;

            // 印刷データを作成します。
            var document = this.GetFixedDocument(originHeight, originWidth);

            // 印刷します。
            var writer = PrintQueue.CreateXpsDocumentWriter(queue);
            writer.Write(document, ticket);

        }

        private FixedDocument GetFixedDocument(double height, double width)
        {
            var document = new FixedDocument();

            // サンプルでは3ページとします。
            for (int i = 0; i < 3; ++i)
            {
                // 1ページに描画する内容を設定します。
                var textBlock = new TextBlock();
                textBlock.Text = $"{i + 1}ページ";

                var canvas = new Canvas();
                Canvas.SetTop(textBlock, height);
                Canvas.SetLeft(textBlock, width);
                canvas.Children.Add(textBlock);

                // 1ページのFixedPageは、PageContentでラップしてFixedDocumentに追加します。
                var page = new FixedPage();
                page.Children.Add(canvas);

                var content = new PageContent();
                content.Child = page;

                // 印刷データにページデータを設定します。
                document.Pages.Add(content);
            }
            return document;
        }


        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            var server = new LocalPrintServer();
            var queue = server.GetPrintQueues().FirstOrDefault(_ => _.Name == "Microsoft Print to PDF");

            // 用紙サイズ(A4)と印刷方向(縦)を設定します。
            var ticket = queue.DefaultPrintTicket;
            ticket.PageMediaSize = new PageMediaSize(PageMediaSizeName.ISOA4);
            ticket.PageOrientation = PageOrientation.Portrait;

            // 余白の上と左位置を取得します。
            var area = queue.GetPrintCapabilities().PageImageableArea;
            var originHeight = area.OriginHeight;
            var originWidth = area.OriginWidth;

            // 印刷データを作成します。
            var document = this.GetFixedDocument(originHeight, originWidth);

            System.IO.MemoryStream streamXPS = new System.IO.MemoryStream();
            using (Package pack = Package.Open(streamXPS, System.IO.FileMode.CreateNew))
            {
                using (var doc = new System.Windows.Xps.Packaging.XpsDocument(pack, CompressionOption.SuperFast))
                {
                    var writer = System.Windows.Xps.Packaging.XpsDocument.CreateXpsDocumentWriter(doc);
                    writer.Write(document, ticket);
                }
            }
            streamXPS.Position = 0;
            var acount = "";
            var filePath = $@"C:\Users\{acount}\Desktop\テスト印刷.pdf";
            XpsPrintHelper.Print(streamXPS, queue.Name, "jobname", false, filePath);
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {

            var server = new LocalPrintServer();
            var queue = server.GetPrintQueues().FirstOrDefault(_ => _.Name == "Microsoft Print to PDF");

            // 用紙サイズ(A4)と印刷方向(縦)を設定します。
            var ticket = queue.DefaultPrintTicket;
            ticket.PageMediaSize = new PageMediaSize(PageMediaSizeName.ISOA4);
            ticket.PageOrientation = PageOrientation.Portrait;

            // 余白の上と左位置を取得します。
            var area = queue.GetPrintCapabilities().PageImageableArea;
            var originHeight = area.OriginHeight;
            var originWidth = area.OriginWidth;

            FixedDocument document = new FixedDocument();

            var p = new TestPage();
            FixedPage fp = CreateFixedPage(p);
            PageContent pc = new PageContent();
            pc.Child = fp;
            document.Pages.Add(pc);

            System.IO.MemoryStream streamXPS = new System.IO.MemoryStream();
            using (Package pack = Package.Open(streamXPS, System.IO.FileMode.CreateNew))
            {
                using (var doc = new System.Windows.Xps.Packaging.XpsDocument(pack, CompressionOption.SuperFast))
                {
                    var writer = System.Windows.Xps.Packaging.XpsDocument.CreateXpsDocumentWriter(doc);
                    writer.Write(document, ticket);
                }
            }
            streamXPS.Position = 0;
            var acount = "";
            var filePath = $@"C:\Users\{acount}\Desktop\テスト印刷2.pdf";
            XpsPrintHelper.Print(streamXPS, queue.Name, "jobname", false, filePath);


        }


        private static FixedPage CreateFixedPage(Page page)
        {
            Action dummy = () => { };
            var disp = System.Windows.Threading.Dispatcher.CurrentDispatcher;
            FixedPage fixpage = new FixedPage();
            {
                Frame fm = new Frame();

                fm.Content = page;
                disp.Invoke(dummy, System.Windows.Threading.DispatcherPriority.Loaded);//バインディングとかいろいろさせる必要がある場合

                FixedPage.SetLeft(fm, 0);
                FixedPage.SetTop(fm, 0);

                fixpage.Children.Add(fm);

                Size sz = new Size(page.Width, page.Height);

                fixpage.Width = page.Width;
                fixpage.Height = page.Height;

                fixpage.Measure(sz);
                fixpage.Arrange(new Rect(new Point(), sz));
                fixpage.UpdateLayout();


            }
            return fixpage;
        }

    }
}
