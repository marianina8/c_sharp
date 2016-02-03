using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Diagnostics;
using System.Net;
using System.Text.RegularExpressions;
using System.IO;
using System.Text;
using HtmlAgilityPack;

namespace ITS
{

    public class Product
    {
        public string Name { get; set; }
        public string Price { get; set; }
    }

    public partial class crawler : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            
        }

        protected void btnStart_Click(object sender, EventArgs e)
        {
            
            // add these to configuration file
            int productLimit = 500; // Limit to how many products exist in a single text file
            string fileLocation = "C:\\\\"; // Location of text files
            string baseFileName = "products";

            // Declare list of products
            List<Product> products = new List<Product>();

            HtmlWeb hw = new HtmlWeb();
            HtmlDocument doc = hw.Load(txtURL.Text);
            //loop through all links found on URL page
            foreach (HtmlNode link in doc.DocumentNode.SelectNodes("//a[@href]"))
            {

                
                // if the link is a product page by containing the start of a product link
                if (link.Name == "a" && link.Attributes["href"].Value.StartsWith("http://www.newegg.com/Product/Product.aspx"))
                {
                    // create Product object
                    Product p = new Product();

                    // grab html from each product page
                    HtmlDocument productPage = hw.Load(link.Attributes["href"].Value);

                    //extract product names
                    p.Name = productPage.DocumentNode.SelectSingleNode(".//span[@itemprop='name']").InnerText;
                    
                    // Need to implement price extraction
                    p.Price = "N/A"; //productPage.DocumentNode.SelectSingleNode("//[@class='price-current'").InnerText;
                    
                    //add product to product list
                    products.Add(p);
                }
            }

            // bind product list to grid
            GridView1.DataSource = products;
            GridView1.DataBind();

            // save to local text file
            int val = products.Count / 500;
            if (val==0)
            {
                // print everything to one file

                TextWriter tw = new StreamWriter(fileLocation+"\\"+baseFileName+".txt");

                foreach (Product p in products)
                    tw.WriteLine(p.Name.Trim()+","+p.Price.Trim());

                tw.Close();
            }
            else if (val > 0)
            {
                // use the baseFileName and create incremented files and check mod for additional products for last file.
            }

        }
    }
}