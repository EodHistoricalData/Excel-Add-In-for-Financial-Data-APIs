using EODAddIn.Program;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

using static EOD.API;


namespace EODAddIn.Forms
{
    public partial class FrmScreener : Form
    {
        public List<(Field, Operation, string)> Filters { get; set; } = new List<(Field, Operation, string)>();
        public List<Signal> Signals { get; set; } = new List<Signal>();
        public (Field, Order)? Sort { get; set; }
        public int Limit { get; set; }

        private Dictionary<string, string> _fields = new Dictionary<string, string>()
        {
            { "market capitalization", "number" },
            { "earnings share", "number" },
            { "dividend yield", "number" },
            { "refund 1d p", "number" },
            { "refund 5d p", "number" },
            { "avgvol 1d", "number" },
            { "avgvol 200d", "number" },
        };
        private List<string> _operation = new List<string>()
        {
            { "=" },
            { ">" },
            { "<" },
            { ">=" },
            { "<=" },
            { "!=" },
        };

        private List<(string, string)> _industriesAndSectors = new List<(string, string)>()
            {
            ( "Agricultural Chemicals", "Basic Materials" ),
            ( "Agricultural Inputs", "Basic Materials"),
            ( "Aluminum", "Basic Materials"),
            ( "Building Materials", "Basic Materials"),
            ( "Chemicals", "Basic Materials"),
            ( "Chemicals - Major Diversified", "Basic Materials"),
            ( "Coal", "Basic Materials"),
            ( "Coking Coal", "Basic Materials"),
            ( "Copper", "Basic Materials"),
            ( "Gold", "Basic Materials"),
            ( "Independent Oil & Gas", "Basic Materials"),
            ( "Industrial Metals & Minerals", "Basic Materials"),
            ( "Lumber & Wood Production", "Basic Materials"),
            ( "Major Integrated Oil & Gas", "Basic Materials"),
            ( "Nonmetallic Mineral Mining", "Basic Materials"),
            ( "Oil & Gas Drilling & Exploration", "Basic Materials"),
            ( "Oil & Gas Equipment & Services", "Basic Materials"),
            ( "Oil & Gas Pipelines", "Basic Materials"),
            ( "Oil & Gas Refining & Marketing", "Basic Materials"),
            ( "Other Industrial Metals & Mining", "Basic Materials"),
            ( "Other Precious Metals & Mining", "Basic Materials"),
            ( "Paper & Paper Products", "Basic Materials"),
            ( "Silver", "Basic Materials"),
            ( "Specialty Chemicals", "Basic Materials"),
            ( "Steel", "Basic Materials"),
            ( "Steel & Iron", "Basic Materials"),
            ( "Synthetics", "Basic Materials"),
            ( "Advertising Agencies", "Communication Services"),
            ( "Broadcasting", "Communication Services"),
            ( "Electronic Gaming & Multimedia", "Communication Services"),
            ( "Entertainment", "Communication Services"),
            ( "Internet Content & Information", "Communication Services"),
            ( "Pay TV", "Communication Services"),
            ( "Publishing", "Communication Services"),
            ( "Telecom Services", "Communication Services"),
            ( "Conglomerates", "Conglomerates"),
            ( "Advertising Agencies", "Consumer Cyclical"),
            ( "Apparel Manufacturing", "Consumer Cyclical"),
            ( "Apparel Retail", "Consumer Cyclical"),
            ( "Apparel Stores", "Consumer Cyclical"),
            ( "Auto & Truck Dealerships", "Consumer Cyclical"),
            ( "Auto Manufacturers", "Consumer Cyclical"),
            ( "Auto Parts", "Consumer Cyclical"),
            ( "Broadcasting - Radio", "Consumer Cyclical"),
            ( "Broadcasting - TV", "Consumer Cyclical"),
            ( "Department Stores", "Consumer Cyclical"),
            ( "Footwear & Accessories", "Consumer Cyclical"),
            ( "Furnishings", "Consumer Cyclical"),
            ( "Gambling", "Consumer Cyclical"),
            ( "Home Furnishings & Fixtures", "Consumer Cyclical"),
            ( "Home Improvement Retail", "Consumer Cyclical"),
            ( "Home Improvement Stores", "Consumer Cyclical"),
            ( "Internet Retail", "Consumer Cyclical"),
            ( "Leisure", "Consumer Cyclical"),
            ( "Lodging", "Consumer Cyclical"),
            ( "Luxury Goods", "Consumer Cyclical"),
            ( "Marketing Services", "Consumer Cyclical"),
            ( "Media - Diversified", "Consumer Cyclical"),
            ( "Packaging & Containers", "Consumer Cyclical"),
            ( "Personal Services", "Consumer Cyclical"),
            ( "Publishing", "Consumer Cyclical"),
            ( "Recreational Vehicles", "Consumer Cyclical"),
            ( "Residential Construction", "Consumer Cyclical"),
            ( "Resorts & Casinos", "Consumer Cyclical"),
            ( "Restaurants", "Consumer Cyclical"),
            ( "Rubber & Plastics", "Consumer Cyclical"),
            ( "Specialty Retail", "Consumer Cyclical"),
            ( "Textile Manufacturing", "Consumer Cyclical"),
            ( "Travel Services", "Consumer Cyclical"),
            ( "Beverages - Brewers", "Consumer Defensive"),
            ( "Beverages - Soft Drinks", "Consumer Defensive"),
            ( "Beverages - Wineries & Distilleries", "Consumer Defensive"),
            ( "Beverages-Brewers", "Consumer Defensive"),
            ( "Beverages-Non-Alcoholic", "Consumer Defensive"),
            ( "Beverages-Wineries & Distilleries", "Consumer Defensive"),
            ( "Confectioners", "Consumer Defensive"),
            ( "Discount Stores", "Consumer Defensive"),
            ( "Education & Training Services", "Consumer Defensive"),
            ( "Farm Products", "Consumer Defensive"),
            ( "Food Distribution", "Consumer Defensive"),
            ( "Grocery Stores", "Consumer Defensive"),
            ( "Household & Personal Products", "Consumer Defensive"),
            ( "Packaged Foods", "Consumer Defensive"),
            ( "Pharmaceutical Retailers", "Consumer Defensive"),
            ( "Tobacco", "Consumer Defensive"),
            ( "Appliances", "Consumer Goods"),
            ( "Auto Manufacturers - Major", "Consumer Goods"),
            ( "Auto Parts", "Consumer Goods"),
            ( "Beverages - Brewers", "Consumer Goods"),
            ( "Beverages - Soft Drinks", "Consumer Goods"),
            ( "Beverages - Wineries & Distillers", "Consumer Goods"),
            ( "Business Equipment", "Consumer Goods"),
            ( "Cigarettes", "Consumer Goods"),
            ( "Cleaning Products", "Consumer Goods"),
            ( "Confectioners", "Consumer Goods"),
            ( "Dairy Products", "Consumer Goods"),
            ( "Electronic Equipment", "Consumer Goods"),
            ( "Farm Products", "Consumer Goods"),
            ( "Food - Major Diversified", "Consumer Goods"),
            ( "Home Furnishings & Fixtures", "Consumer Goods"),
            ( "Housewares & Accessories", "Consumer Goods"),
            ( "Meat Products", "Consumer Goods"),
            ( "Office Supplies", "Consumer Goods"),
            ( "Packaging & Containers", "Consumer Goods"),
            ( "Paper & Paper Products", "Consumer Goods"),
            ( "Personal Products", "Consumer Goods"),
            ( "Photographic Equipment & Supplies", "Consumer Goods"),
            ( "Processed & Packaged Goods", "Consumer Goods"),
            ( "Recreational Goods", "Consumer Goods"),
            ( "Recreational Vehicles", "Consumer Goods"),
            ( "REIT - Retail", "Consumer Goods"),
            ( "Rubber & Plastics", "Consumer Goods"),
            ( "Sporting Goods", "Consumer Goods"),
            ( "Textile - Apparel Clothing", "Consumer Goods"),
            ( "Textile - Apparel Footwear & Accessories", "Consumer Goods"),
            ( "Tobacco Products", "Consumer Goods"),
            ( "Toys & Games", "Consumer Goods"),
            ( "Trucks & Other Vehicles", "Consumer Goods"),
            ( "Oil & Gas Drilling", "Energy"),
            ( "Oil & Gas E&P", "Energy"),
            ( "Oil & Gas Equipment & Services", "Energy"),
            ( "Oil & Gas Integrated", "Energy"),
            ( "Oil & Gas Midstream", "Energy"),
            ( "Oil & Gas Refining & Marketing", "Energy"),
            ( "Thermal Coal", "Energy"),
            ( "Uranium", "Energy"),
            ( "Accident & Health Insurance", "Financial"),
            ( "Asset Management", "Financial"),
            ( "Closed-End Fund - Debt", "Financial"),
            ( "Closed-End Fund - Equity", "Financial"),
            ( "Closed-End Fund - Foreign", "Financial"),
            ( "Credit Services", "Financial"),
            ( "Diversified Investments", "Financial"),
            ( "Foreign Money Center Banks", "Financial"),
            ( "Foreign Regional Banks", "Financial"),
            ( "Insurance Brokers", "Financial"),
            ( "Investment Brokerage - National", "Financial"),
            ( "Investment Brokerage - Regional", "Financial"),
            ( "Life Insurance", "Financial"),
            ( "Money Center Banks", "Financial"),
            ( "Mortgage Investment", "Financial"),
            ( "Property & Casualty Insurance", "Financial"),
            ( "Property Management", "Financial"),
            ( "Real Estate Development", "Financial"),
            ( "Regional - Mid-Atlantic Banks", "Financial"),
            ( "Regional - Midwest Banks", "Financial"),
            ( "Regional - Northeast Banks", "Financial"),
            ( "Regional - Pacific Banks", "Financial"),
            ( "Regional - Southeast Banks", "Financial"),
            ( "Regional - Southwest Banks", "Financial"),
            ( "REIT - Diversified", "Financial"),
            ( "REIT - Healthcare Facilities", "Financial"),
            ( "REIT - Hotel/Motel", "Financial"),
            ( "REIT - Industrial", "Financial"),
            ( "REIT - Office", "Financial"),
            ( "REIT - Residential", "Financial"),
            ( "REIT - Retail", "Financial"),
            ( "Savings & Loans", "Financial"),
            ( "Surety & Title Insurance", "Financial"),
            ( "Asset Management", "Financial Services"),
            ( "Banks - Global", "Financial Services"),
            ( "Banks - Regional - Africa", "Financial Services"),
            ( "Banks - Regional - Asia", "Financial Services"),
            ( "Banks - Regional - Australia", "Financial Services"),
            ( "Banks - Regional - Canada", "Financial Services"),
            ( "Banks - Regional - Europe", "Financial Services"),
            ( "Banks - Regional - Latin America", "Financial Services"),
            ( "Banks - Regional - US", "Financial Services"),
            ( "Banks-Diversified", "Financial Services"),
            ( "Banks-Regional", "Financial Services"),
            ( "Capital Markets", "Financial Services"),
            ( "Credit Services", "Financial Services"),
            ( "Financial Conglomerates", "Financial Services"),
            ( "Financial Data & Stock Exchanges", "Financial Services"),
            ( "Financial Exchanges", "Financial Services"),
            ( "Insurance - Diversified", "Financial Services"),
            ( "Insurance - Life", "Financial Services"),
            ( "Insurance - Property & Casualty", "Financial Services"),
            ( "Insurance - Reinsurance", "Financial Services"),
            ( "Insurance - Specialty", "Financial Services"),
            ( "Insurance Brokers", "Financial Services"),
            ( "Insurance-Diversified", "Financial Services"),
            ( "Insurance-Life", "Financial Services"),
            ( "Insurance-Property & Casualty", "Financial Services"),
            ( "Insurance-Reinsurance", "Financial Services"),
            ( "Insurance-Specialty", "Financial Services"),
            ( "Mortgage Finance", "Financial Services"),
            ( "Savings & Cooperative Banks", "Financial Services"),
            ( "Shell Companies", "Financial Services"),
            ( "Specialty Finance", "Financial Services"),
            ( "Biotechnology", "Healthcare"),
            ( "Diagnostic Substances", "Healthcare"),
            ( "Diagnostics & Research", "Healthcare"),
            ( "Drug Delivery", "Healthcare"),
            ( "Drug Manufacturers - Major", "Healthcare"),
            ( "Drug Manufacturers - Other", "Healthcare"),
            ( "Drug Manufacturers - Specialty & Generic", "Healthcare"),
            ( "Drug Manufacturers-General", "Healthcare"),
            ( "Drug Manufacturers-Specialty & Generic", "Healthcare"),
            ( "Drug Related Products", "Healthcare"),
            ( "Drugs - Generic", "Healthcare"),
            ( "Health Care Plans", "Healthcare"),
            ( "Health Information Services", "Healthcare"),
            ( "Healthcare Plans", "Healthcare"),
            ( "Home Health Care", "Healthcare"),
            ( "Hospitals", "Healthcare"),
            ( "Long-Term Care Facilities", "Healthcare"),
            ( "Medical Appliances & Equipment", "Healthcare"),
            ( "Medical Care", "Healthcare"),
            ( "Medical Care Facilities", "Healthcare"),
            ( "Medical Devices", "Healthcare"),
            ( "Medical Distribution", "Healthcare"),
            ( "Medical Instruments & Supplies", "Healthcare"),
            ( "Medical Laboratories & Research", "Healthcare"),
            ( "Medical Practitioners", "Healthcare"),
            ( "Pharmaceutical Retailers", "Healthcare"),
            ( "Specialized Health Services", "Healthcare"),
            ( "Aerospace/Defense - Major Diversified", "Industrial Goods"),
            ( "Aerospace/Defense Products & Services", "Industrial Goods"),
            ( "Cement", "Industrial Goods"),
            ( "Diversified Machinery", "Industrial Goods"),
            ( "Farm & Construction Machinery", "Industrial Goods"),
            ( "General Building Materials", "Industrial Goods"),
            ( "General Contractors", "Industrial Goods"),
            ( "Heavy Construction", "Industrial Goods"),
            ( "Industrial Electrical Equipment", "Industrial Goods"),
            ( "Industrial Equipment & Components", "Industrial Goods"),
            ( "Lumber", "Industrial Goods"),
            ( "Machine Tools & Accessories", "Industrial Goods"),
            ( "Manufactured Housing", "Industrial Goods"),
            ( "Metal Fabrication", "Industrial Goods"),
            ( "Pollution & Treatment Controls", "Industrial Goods"),
            ( "Residential Construction", "Industrial Goods"),
            ( "Small Tools & Accessories", "Industrial Goods"),
            ( "Textile Industrial", "Industrial Goods"),
            ( "Waste Management", "Industrial Goods"),
            ( "Aerospace & Defense", "Industrials"),
            ( "Airlines", "Industrials"),
            ( "Airports & Air Services", "Industrials"),
            ( "Building Products & Equipment", "Industrials"),
            ( "Business Equipment", "Industrials"),
            ( "Business Equipment & Supplies", "Industrials"),
            ( "Business Services", "Industrials"),
            ( "Conglomerates", "Industrials"),
            ( "Consulting Services", "Industrials"),
            ( "Diversified Industrials", "Industrials"),
            ( "Electrical Equipment & Parts", "Industrials"),
            ( "Engineering & Construction", "Industrials"),
            ( "Farm & Construction Equipment", "Industrials"),
            ( "Farm & Heavy Construction Machinery", "Industrials"),
            ( "Industrial Distribution", "Industrials"),
            ( "Infrastructure Operations", "Industrials"),
            ( "Integrated Freight & Logistics", "Industrials"),
            ( "Integrated Shipping & Logistics", "Industrials"),
            ( "Marine Shipping", "Industrials"),
            ( "Metal Fabrication", "Industrials"),
            ( "Pollution & Treatment Controls", "Industrials"),
            ( "Railroads", "Industrials"),
            ( "Rental & Leasing Services", "Industrials"),
            ( "Security & Protection Services", "Industrials"),
            ( "Shipping & Ports", "Industrials"),
            ( "Specialty Business Services", "Industrials"),
            ( "Specialty Industrial Machinery", "Industrials"),
            ( "Staffing & Employment Services", "Industrials"),
            ( "Staffing & Outsourcing Services", "Industrials"),
            ( "Tools & Accessories", "Industrials"),
            ( "Truck Manufacturing", "Industrials"),
            ( "Trucking", "Industrials"),
            ( "Waste Management", "Industrials"),
            ( "Other", "Other"),
            ( "Real Estate - General", "Real Estate"),
            ( "Real Estate Services", "Real Estate"),
            ( "Real Estate-Development", "Real Estate"),
            ( "Real Estate-Diversified", "Real Estate"),
            ( "REIT - Diversified", "Real Estate"),
            ( "REIT - Healthcare Facilities", "Real Estate"),
            ( "REIT - Hotel & Motel", "Real Estate"),
            ( "REIT - Industrial", "Real Estate"),
            ( "REIT - Office", "Real Estate"),
            ( "REIT - Residential", "Real Estate"),
            ( "REIT - Retail", "Real Estate"),
            ( "REIT-Diversified", "Real Estate"),
            ( "REIT-Healthcare Facilities", "Real Estate"),
            ( "REIT-Hotel & Motel", "Real Estate"),
            ( "REIT-Industrial", "Real Estate"),
            ( "REIT-Mortgage", "Real Estate"),
            ( "REIT-Office", "Real Estate"),
            ( "REIT-Residential", "Real Estate"),
            ( "REIT-Retail", "Real Estate"),
            ( "REIT-Specialty", "Real Estate"),
            ( "Advertising Agencies", "Services"),
            ( "Air Delivery & Freight Services", "Services"),
            ( "Air Services", "Services"),
            ( "Apparel Stores", "Services"),
            ( "Auto Dealerships", "Services"),
            ( "Auto Parts Stores", "Services"),
            ( "Auto Parts Wholesale", "Services"),
            ( "Basic Materials Wholesale", "Services"),
            ( "Broadcasting - Radio", "Services"),
            ( "Broadcasting - TV", "Services"),
            ( "Building Materials Wholesale", "Services"),
            ( "Business Services", "Services"),
            ( "Catalog & Mail Order Houses", "Services"),
            ( "CATV Systems", "Services"),
            ( "Computers Wholesale", "Services"),
            ( "Consumer Services", "Services"),
            ( "Department Stores", "Services"),
            ( "Discount", "Services"),
            ( "Drug Stores", "Services"),
            ( "Drugs Wholesale", "Services"),
            ( "Education & Training Services", "Services"),
            ( "Electronics Stores", "Services"),
            ( "Electronics Wholesale", "Services"),
            ( "Entertainment - Diversified", "Services"),
            ( "Food Wholesale", "Services"),
            ( "Gaming Activities", "Services"),
            ( "General Entertainment", "Services"),
            ( "Grocery Stores", "Services"),
            ( "Home Furnishing Stores", "Services"),
            ( "Home Improvement Stores", "Services"),
            ( "Industrial Equipment Wholesale", "Services"),
            ( "Information Technology Services", "Services"),
            ( "Jewelry Stores", "Services"),
            ( "Lodging", "Services"),
            ( "Major Airlines", "Services"),
            ( "Management Services", "Services"),
            ( "Marketing Services", "Services"),
            ( "Medical Equipment Wholesale", "Services"),
            ( "Movie Production", "Services"),
            ( "Music & Video Stores", "Services"),
            ( "Personal Services", "Services"),
            ( "Publishing - Books", "Services"),
            ( "Publishing - Newspapers", "Services"),
            ( "Publishing - Periodicals", "Services"),
            ( "Railroads", "Services"),
            ( "Regional Airlines", "Services"),
            ( "Rental & Leasing Services", "Services"),
            ( "Research Services", "Services"),
            ( "Resorts & Casinos", "Services"),
            ( "Restaurants", "Services"),
            ( "Security & Protection Services", "Services"),
            ( "Shipping", "Services"),
            ( "Specialty Eateries", "Services"),
            ( "Specialty Retail", "Services"),
            ( "Sporting Activities", "Services"),
            ( "Sporting Goods Stores", "Services"),
            ( "Staffing & Outsourcing Services", "Services"),
            ( "Technical Services", "Services"),
            ( "Toy & Hobby Stores", "Services"),
            ( "Trucking", "Services"),
            ( "Wholesale", "Services"),
            ( "Application Software", "Technology"),
            ( "Business Software & Services", "Technology"),
            ( "Communication Equipment", "Technology"),
            ( "Computer Based Systems", "Technology"),
            ( "Computer Distribution", "Technology"),
            ( "Computer Hardware", "Technology"),
            ( "Computer Peripherals", "Technology"),
            ( "Computer Systems", "Technology"),
            ( "Consumer Electronics", "Technology"),
            ( "Contract Manufacturers", "Technology"),
            ( "Data Storage", "Technology"),
            ( "Data Storage Devices", "Technology"),
            ( "Diversified Communication Services", "Technology"),
            ( "Diversified Computer Systems", "Technology"),
            ( "Diversified Electronics", "Technology"),
            ( "Electronic Components", "Technology"),
            ( "Electronic Gaming & Multimedia", "Technology"),
            ( "Electronics & Computer Distribution", "Technology"),
            ( "Electronics Distribution", "Technology"),
            ( "Health Information Services", "Technology"),
            ( "Healthcare Information Services", "Technology"),
            ( "Information & Delivery Services", "Technology"),
            ( "Information Technology Services", "Technology"),
            ( "Internet Content & Information", "Technology"),
            ( "Internet Information Providers", "Technology"),
            ( "Internet Service Providers", "Technology"),
            ( "Internet Software & Services", "Technology"),
            ( "Long Distance Carriers", "Technology"),
            ( "Multimedia & Graphics Software", "Technology"),
            ( "Networking & Communication Devices", "Technology"),
            ( "Personal Computers", "Technology"),
            ( "Printed Circuit Boards", "Technology"),
            ( "Processing Systems & Products", "Technology"),
            ( "Scientific & Technical Instruments", "Technology"),
            ( "Security Software & Services", "Technology"),
            ( "Semiconductor - Broad Line", "Technology"),
            ( "Semiconductor - Integrated Circuits", "Technology"),
            ( "Semiconductor - Specialized", "Technology"),
            ( "Semiconductor Equipment & Materials", "Technology"),
            ( "Semiconductor Memory", "Technology"),
            ( "Semiconductor- Memory Chips", "Technology"),
            ( "Semiconductors", "Technology"),
            ( "Software - Application", "Technology"),
            ( "Software - Infrastructure", "Technology"),
            ( "Software-Application", "Technology"),
            ( "Software-Infrastructure", "Technology"),
            ( "Solar", "Technology"),
            ( "Technical & System Software", "Technology"),
            ( "Telecom Services - Domestic", "Technology"),
            ( "Telecom Services - Foreign", "Technology"),
            ( "Wireless Communications", "Technology"),
            ( "Diversified Utilities", "Utilities"),
            ( "Electric Utilities", "Utilities"),
            ( "Foreign Utilities", "Utilities"),
            ( "Gas Utilities", "Utilities"),
            ( "Utilities - Diversified", "Utilities"),
            ( "Utilities - Independent Power Producers", "Utilities"),
            ( "Utilities - Regulated Electric", "Utilities"),
            ( "Utilities - Regulated Gas", "Utilities"),
            ( "Utilities - Regulated Water", "Utilities"),
            ( "Utilities-Diversified", "Utilities"),
            ( "Utilities-Independent Power Producers", "Utilities"),
            ( "Utilities-Regulated Electric", "Utilities"),
            ( "Utilities-Regulated Gas", "Utilities"),
            ( "Utilities-Regulated Water", "Utilities"),
            ( "Utilities-Renewable", "Utilities"),
            ( "Water Utilities", "Utilities"),


        };

        public FrmScreener()
        {
            InitializeComponent();
            cboSector.Text = Settings.SettingsFields.ScreenerSector;
            cboIndustry.Text = Settings.SettingsFields.ScreenerIndustry;
            txtCode.Text = Settings.SettingsFields.ScreenerCode;
            txtName.Text = Settings.SettingsFields.ScreenerName;
            txtExchange.Text = Settings.SettingsFields.ScreenerExchange;
            numLimit.Value = Settings.SettingsFields.ScreenerLimit;
            chk50d_new_hi.CheckState = Settings.SettingsFields.Screener50d_New_Hi;
            chk50d_new_lo.CheckState = Settings.SettingsFields.Screener50d_New_Lo;
            chk200d_new_hi.CheckState = Settings.SettingsFields.Screener200d_New_Hi;
            chk200d_new_lo.CheckState = Settings.SettingsFields.Screener200d_New_Lo;
            chkWallstreet_hi.CheckState = Settings.SettingsFields.ScreenerWallStreet_Hi;
            chkWallstreet_lo.CheckState = Settings.SettingsFields.ScreenerWallStreet_Lo;
            rbtnSortAsc.Checked = Settings.SettingsFields.ScreenerRbtnSortAsc;
            rbtnSortDesc.Checked = Settings.SettingsFields.ScreenerRbtnSortDesc;
            int dataGridRow=0;
            foreach ((string,string,string) values in Settings.SettingsFields.ScreenerDataGridViewFilters)
            {

                dataGridViewFilters.Rows.Add();
                dataGridViewFilters.Rows[dataGridRow].Cells[0].Value = values.Item1;
                dataGridViewFilters.Rows[dataGridRow].Cells[1].Value = values.Item2;
                dataGridViewFilters.Rows[dataGridRow].Cells[2].Value = values.Item3;
                dataGridRow++;
            }
            Settings.SettingsFields.ScreenerDataGridViewFilters.Clear();
        }
        private void dataGridViewFilters_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            DataGridViewComboBoxCell cell = (DataGridViewComboBoxCell)dataGridViewFilters.Rows[e.RowIndex].Cells[colField.Index];
            foreach (var item in _fields)
            {
                cell.Items.Add(item.Key);
            }
        }

        private void dataGridViewFilters_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == colField.Index && e.RowIndex > -1)
            {
                string val = dataGridViewFilters.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                if (_fields.ContainsKey(val))
                {
                    DataGridViewComboBoxCell cell;
                    List<string> lst;
                    //if (_fields[val] == "string")
                    //{
                    //    cell = (DataGridViewComboBoxCell)dataGridViewFilters.Rows[e.RowIndex].Cells[colOperation.Index];
                    //    lst = _operationString;
                    //}
                    //else
                    //{
                        cell = (DataGridViewComboBoxCell)dataGridViewFilters.Rows[e.RowIndex].Cells[colOperation.Index];
                        lst = _operation;
                    //}

                    cell.Items.Clear();
                    foreach (var item in lst)
                    {
                        cell.Items.Add(item);
                    }

                    if (cell.Items.Count == 1)
                    {
                        cell.Value = cell.Items[0].ToString();
                    }
                }
            }
        }

        private void btnAddFilter_Click(object sender, EventArgs e)
        {
            dataGridViewFilters.Rows.Add();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            try
            {
                SetFilteres();
                if (Filters.Count==0)
                {
                    return; 
                }
                SetSignals();
                SetSort();
                Limit = (int)numLimit.Value+1;
                DialogResult = DialogResult.OK;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Screener error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            Settings.SettingsFields.ScreenerSector = cboSector.Text;
            Settings.SettingsFields.ScreenerIndustry = cboIndustry.Text;
            Settings.SettingsFields.ScreenerCode = txtCode.Text;
            Settings.SettingsFields.ScreenerName = txtName.Text;
            Settings.SettingsFields.ScreenerExchange = txtExchange.Text;
            Settings.SettingsFields.ScreenerLimit = (int)numLimit.Value;
            Settings.SettingsFields.Screener50d_New_Lo = chk50d_new_lo.CheckState;
            Settings.SettingsFields.Screener50d_New_Hi = chk50d_new_hi.CheckState;
            Settings.SettingsFields.Screener200d_New_Hi = chk200d_new_hi.CheckState;
            Settings.SettingsFields.Screener200d_New_Lo = chk200d_new_lo.CheckState;
            Settings.SettingsFields.ScreenerBookValue_Neg = chkBookvalue_neg.CheckState;
            Settings.SettingsFields.ScreenerBookValue_Pos = chkBookvalue_pos.CheckState;
            Settings.SettingsFields.ScreenerWallStreet_Lo = chkWallstreet_lo.CheckState;
            Settings.SettingsFields.ScreenerWallStreet_Hi = chkWallstreet_hi.CheckState;
            Settings.SettingsFields.ScreenerRbtnSortAsc = rbtnSortAsc.Checked;
            Settings.SettingsFields.ScreenerRbtnSortDesc = rbtnSortDesc.Checked;
            for(int i = 0; i < dataGridViewFilters.Rows.Count; i++)
            {
                string field = dataGridViewFilters.Rows[i].Cells[0].Value.ToString() ;
                string operation = dataGridViewFilters.Rows[i].Cells[1].Value.ToString();
                string value = dataGridViewFilters.Rows[i].Cells[2].Value.ToString();
                Settings.SettingsFields.ScreenerDataGridViewFilters.Add((field,operation,value));
            }
            Settings.Save();
        }

        private void SetFilteres()
        {
            Filters.Clear();
            foreach (DataGridViewRow row in dataGridViewFilters.Rows)
            {
                if (row.Cells[colField.Index].Value == null) continue;
                if (row.Cells[colOperation.Index].Value == null) throw new Exception("Select operation type from the list");
                if (row.Cells[colValue.Index].Value == null) throw new Exception("Select a value for the operation");

                Field field;
                Operation operation;
                switch (row.Cells[colField.Index].Value.ToString())
                {            
                    case "market capitalization":
                        field = Field.MarketCapitalization;
                        break;
                    case "earnings share":
                        field = Field.EarningsShare;
                        break;
                    case "dividend yield":
                        field = Field.DividendYield;
                        break;
                    case "refund 1d p":
                        field = Field.Refund1dP;
                        break;
                    case "refund 5d p":
                        field = Field.Refund5dP;
                        break;
                    case "avgvol 1d":
                        field = Field.Refund5dP;
                        break;
                    case "avgvol 200d":
                        field = Field.Refund5dP;
                        break;

                    default:
                        throw new Exception("Select a field");
                }

                switch (row.Cells[colOperation.Index].Value.ToString())
                {
                    case "=":
                        operation = Operation.Equals;
                        break;
                    case ">":
                        operation = Operation.More;
                        break;
                    case "<":
                        operation = Operation.Less;
                        break;
                    case ">=":
                        operation = Operation.NotLess;
                        break;
                    case "<=":
                        operation = Operation.NotMore;
                        break;
                    case "!=":
                        operation = Operation.NotEquals;
                        break;

                    default:
                        throw new Exception("Select a operation");
                }

                (Field, Operation, string) filter = (field, operation, row.Cells[colValue.Index].Value.ToString());
                Filters.Add(filter);
            }

            if (!string.IsNullOrEmpty(txtCode.Text))
            {
                (Field, Operation, string) filter = (Field.Code, Operation.Equals, txtCode.Text);
                Filters.Add(filter);
            }

            if (!string.IsNullOrEmpty(txtName.Text))
            {
                (Field, Operation, string) filter = (Field.Name, Operation.Equals, txtName.Text);
                Filters.Add(filter);
            }
            if (!string.IsNullOrEmpty(txtExchange.Text))
            {
                (Field, Operation, string) filter = (Field.Exchange, Operation.Equals, txtExchange.Text);
                Filters.Add(filter);
            }
            if (!string.IsNullOrEmpty(cboSector.Text))
            {
                (Field, Operation, string) filter = (Field.Sector, Operation.Equals, cboSector.Text);
                Filters.Add(filter);
            }
            if (!string.IsNullOrEmpty(cboIndustry.Text))
            {
                (Field, Operation, string) filter = (Field.Industry, Operation.Equals, cboIndustry.Text);
                Filters.Add(filter);
            }
            if (Filters.Count == 0)
            {
                MessageBox.Show("Not enough filters", "Error",  MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void SetSignals()
        {
            Signals = new List<Signal>();
            if (chk50d_new_lo.Checked) Signals.Add(Signal.New_50d_low);
            if (chk50d_new_hi.Checked) Signals.Add(Signal.New_50d_hi);
            if (chk200d_new_lo.Checked) Signals.Add(Signal.New_200d_low);
            if (chk200d_new_hi.Checked) Signals.Add(Signal.New_200d_hi);
            if (chkBookvalue_neg.Checked) Signals.Add(Signal.Bookvalue_neg);
            if (chkBookvalue_pos.Checked) Signals.Add(Signal.Bookvalue_pos);
            if (chkWallstreet_lo.Checked) Signals.Add(Signal.Wallstreet_low);
            if (chkWallstreet_hi.Checked) Signals.Add(Signal.Wallstreet_hi);

            if (Signals.Count == 0) Signals = null;
        }

        private void SetSort()
        {

            if (cboSortField.SelectedIndex == -1) return;
            Field field;
            Order order;

            switch (cboSortField.Text)
            {
                case "code":
                    field = Field.Code;
                    break;

                case "name":
                    field = Field.Name;
                    break;
                case "exchange":
                    field = Field.Exchange;
                    break;
                case "sector":
                    field = Field.Sector;
                    break;
                case "industry":
                    field = Field.Industry;
                    break;
                case "market_capitalization":
                    field = Field.MarketCapitalization;
                    break;
                case "earnings_share":
                    field = Field.EarningsShare;
                    break;
                case "dividend_yield":
                    field = Field.DividendYield;
                    break;
                case "refund_1d_p":
                    field = Field.Refund1dP;
                    break;
                case "refund_5d_p":
                    field = Field.Refund5dP;
                    break;
                case "avgvol_1d":
                    field = Field.Refund5dP;
                    break;
                case "avgvol_200d":
                    field = Field.Refund5dP;
                    break;

                default:
                    Sort = null;
                    return;
            }

            if (rbtnSortAsc.Checked)
            {
                order = Order.Ascending;
            }
            else
            {
                order = Order.Descending;
            }

            Sort = (field, order);
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void FrmScreener_Load(object sender, EventArgs e)
        {
            List<string> list = new List<string>();
            foreach (var item in _industriesAndSectors)
            {
                if (!list.Contains(item.Item2))
                {
                    list.Add(item.Item2);
                }
            }
            cboSector.Items.AddRange(list.Distinct().ToArray());
        }

        private void cboSector_SelectedIndexChanged(object sender, EventArgs e)
        {
            cboIndustry.Items.Clear();
            List<string> list = new List<string>();

            foreach (var item in _industriesAndSectors)
            {
                if (item.Item2==cboSector.Text)
                {
                    list.Add(item.Item1);
                }
            }
            cboIndustry.Items.AddRange(list.Distinct().ToArray());
        }
    }
}
