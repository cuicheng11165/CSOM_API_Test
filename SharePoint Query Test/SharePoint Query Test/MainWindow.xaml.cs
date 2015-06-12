using Microsoft.CSharp.RuntimeBinder;
using Microsoft.SharePoint.Client;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Dynamic;
using System.Linq;
using System.Net;
using System.Runtime.CompilerServices;
using System.Security;
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

namespace SharePoint_Query_Test
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
        ClientContext context;

        Microsoft.SharePoint.Client.List list;

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string webUrl = SiteCollectionUrlText.Text.Trim();

                context = new ClientContext(webUrl);

                if (string.Equals(this.CredentialType.Text, "", StringComparison.OrdinalIgnoreCase))
                {
                    context.Credentials = new NetworkCredential(this.UserNameTextBox.Text.Trim(), this.PasswordTextBox.Text);
                }
                else if (string.Equals(this.CredentialType.Text, "", StringComparison.OrdinalIgnoreCase))
                {
                    SecureString secureStr = new SecureString();
                    foreach (var c in this.PasswordTextBox.Text.Trim())
                    {
                        secureStr.AppendChar(c);
                    }
                    context.Credentials = new SharePointOnlineCredentials(this.UserNameTextBox.Text.Trim(), secureStr);
                }

                list = context.Web.Lists.GetByTitle(this.ListTitleTextBox.Text.Trim());

                var query = new CamlQuery();

                var document = this.QueryText.Document;

                var range = new TextRange(document.ContentStart, document.ContentEnd);


                query.ViewXml = range.Text;

                var items = list.GetItems(query);

                context.Load(items);

                context.ExecuteQuery();


                List<Result> resultList = new List<Result>();

                foreach (var item in items)
                {
                    resultList.Add(new Result
                    {
                        FileRef = item.FieldValues["FileRef"].ToString(),
                        FileDirRef = item.FieldValues["FileDirRef"].ToString(),
                        ID = item.FieldValues["ID"].ToString(),
                    });
                }



                this.ResultDataGrid.DataContext = resultList;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }

        private void QueryText_TextChanged(object sender, TextChangedEventArgs e)
        {
            this.ResultDataGrid.DataContext = null;
        }
    }

    public class Result
    {
        public string FileRef { set; get; }

        public string FileDirRef { set; get; }

        public string ID { set; get; }



    }


}
