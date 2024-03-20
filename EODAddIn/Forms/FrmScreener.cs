using EODAddIn.BL.Screener;

using Microsoft.Office.Core;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

using static EOD.API;


namespace EODAddIn.Forms
{
    public partial class FrmScreener : Form
    {
        public Screener Screener { get; private set; }

        private readonly Dictionary<string, string> _fields = new Dictionary<string, string>()
        {
            { "market capitalization", "number" },
            { "earnings share", "number" },
            { "dividend yield", "number" },
            { "refund 1d p", "number" },
            { "refund 5d p", "number" },
            { "avgvol 1d", "number" },
            { "avgvol 200d", "number" },
        };
        private readonly List<string> _operation = new List<string>()
        {
            { "=" },
            { ">" },
            { "<" },
            { ">=" },
            { "<=" },
            { "!=" },
        };
        private readonly List<(string, string)> _industriesAndSectors = new List<(string, string)>()
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
            ( "Aerospace & Defense", "Industrials"),
            ( "Airlines", "Industrials"),
            ( "Airports & Air Services", "Industrials"),
            ( "Building Products & Equipment", "Industrials"),
            ( "Business Equipment", "Industrials"),
            ( "Business Equipment & Supplies", "Industrials"),
            ( "Business Services", "Industrials"),
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

        public FrmScreener(Screener screener = null)
        {
            Screener = screener;
            InitializeComponent();
            SetForm();
        }

        /// <summary>
        /// Fill form from screener data
        /// </summary>
        private void SetForm()
        {
            if (Screener == null)
            {
                Screener = new Screener
                {
                    NameScreener = "New screener"
                };
            }

            txtNameScreener.Text = Screener.NameScreener;
            cboSector.Text = Screener.Sector;
            cboSector_SelectedIndexChanged(null, null);
            cboIndustry.Text = Screener.Industry;
            txtCode.Text = Screener.Code;
            txtName.Text = Screener.Name;
            txtExchange.Text = Screener.Exchange;
            numLimit.Value = Screener.Limit;

            if (Screener.Signals != null)
            {
                foreach (var signal in Screener.Signals)
                {
                    switch (signal)
                    {
                        case Signal.New_50d_low:
                            break;
                        case Signal.New_50d_hi:
                            break;
                        case Signal.New_200d_low:
                            chk200d_new_lo.Checked = true;
                            break;
                        case Signal.New_200d_hi:
                            chk200d_new_hi.Checked = true;
                            break;
                        case Signal.Bookvalue_neg:
                            chkBookvalue_neg.Checked = true;
                            break;
                        case Signal.Bookvalue_pos:
                            chkBookvalue_pos.Checked = true;
                            break;
                        case Signal.Wallstreet_low:
                            chkWallstreet_lo.Checked = true;
                            break;
                        case Signal.Wallstreet_hi:
                            chkWallstreet_hi.Checked = true;
                            break;
                        default:
                            break;
                    }
                }
            }


            if (Screener.Sort != null)
            {
                if (Screener.Sort.Value.Item2 == Order.Ascending)
                {
                    rbtnSortAsc.Checked = true;
                }
                else
                {
                    rbtnSortDesc.Checked = true;
                }

                switch (Screener.Sort.Value.Item1)
                {
                    case Field.Code:
                        cboSortField.Text = "code";
                        break;
                    case Field.Name:
                        cboSortField.Text = "name";
                        break;
                    case Field.Exchange:
                        cboSortField.Text = "exchange";
                        break;
                    case Field.Sector:
                        cboSortField.Text = "sector";
                        break;
                    case Field.Industry:
                        cboSortField.Text = "industry";
                        break;
                    case Field.MarketCapitalization:
                        cboSortField.Text = "market capitalization";
                        break;
                    case Field.EarningsShare:
                        cboSortField.Text = "earnings share";
                        break;
                    case Field.DividendYield:
                        cboSortField.Text = "dividend yield";
                        break;
                    case Field.Refund1dP:
                        cboSortField.Text = "refund 1d";
                        break;
                    //case Field.Refund5dP:
                    //    cboSortField.Text = "code";
                    //    break;
                    //case Field.Avgvol1d:
                    //    cboSortField.Text = "code";
                    //    break;
                    //case Field.Avgvol200d:
                    //    cboSortField.Text = "code";
                    //    break;
                    default:
                        cboSortField.SelectedIndex = -1;
                        break;
                }

                int dataGridRow = 0;
                foreach (var filter in Screener.Filters)
                {
                    switch (filter.Field)
                    {
                        case Field.Code:
                            continue;
                        case Field.Name:
                            continue;
                        case Field.Exchange:
                            continue;
                        case Field.Sector:
                            continue;
                        case Field.Industry:
                            continue;
                        case Field.MarketCapitalization:

                            dataGridViewFilters.Rows.Add();
                            dataGridViewFilters.Rows[dataGridRow].Cells[0].Value = "market capitalization";

                            break;
                        case Field.EarningsShare:

                            dataGridViewFilters.Rows.Add();
                            dataGridViewFilters.Rows[dataGridRow].Cells[0].Value = "earnings share";

                            break;
                        case Field.DividendYield:

                            dataGridViewFilters.Rows.Add();
                            dataGridViewFilters.Rows[dataGridRow].Cells[0].Value = "dividend yield";

                            break;
                        case Field.Refund1dP:

                            dataGridViewFilters.Rows.Add();
                            dataGridViewFilters.Rows[dataGridRow].Cells[0].Value = "refund 1d p";
                            break;
                        case Field.Refund5dP:

                            dataGridViewFilters.Rows.Add();
                            dataGridViewFilters.Rows[dataGridRow].Cells[0].Value = "refund 5d p";
                            continue;
                        case Field.Avgvol1d:

                            dataGridViewFilters.Rows.Add();
                            dataGridViewFilters.Rows[dataGridRow].Cells[0].Value = "avgvol 1d";

                            continue;
                        case Field.Avgvol200d:

                            dataGridViewFilters.Rows.Add();
                            dataGridViewFilters.Rows[dataGridRow].Cells[0].Value = "avgvol 200d";
                            continue;
                        default:
                            continue;
                    }




                    switch (filter.Operation)
                    {
                        case Operation.Matches:
                            dataGridViewFilters.Rows[dataGridRow].Cells[1].Value = "=";
                            break;
                        case Operation.Equals:
                            dataGridViewFilters.Rows[dataGridRow].Cells[1].Value = "=";
                            break;
                        case Operation.More:
                            dataGridViewFilters.Rows[dataGridRow].Cells[1].Value = ">";
                            break;
                        case Operation.Less:
                            dataGridViewFilters.Rows[dataGridRow].Cells[1].Value = "<";
                            break;
                        case Operation.NotLess:
                            dataGridViewFilters.Rows[dataGridRow].Cells[1].Value = ">=";
                            break;
                        case Operation.NotMore:
                            dataGridViewFilters.Rows[dataGridRow].Cells[1].Value = "<=";
                            break;
                        case Operation.NotEquals:
                            dataGridViewFilters.Rows[dataGridRow].Cells[1].Value = "!=";
                            break;
                        default:
                            break;
                    }



                    dataGridViewFilters.Rows[dataGridRow].Cells[2].Value = filter.Value;
                    dataGridRow++;


                }
            }
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

                    cell = (DataGridViewComboBoxCell)dataGridViewFilters.Rows[e.RowIndex].Cells[colOperation.Index];
                    lst = _operation;

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

        private void Ok_Click(object sender, EventArgs e)
        {
            try
            {
                Screener.NameScreener = txtNameScreener.Text;
                Screener.Sector = cboSector.Text;
                Screener.Industry = cboIndustry.Text;
                Screener.Code = txtCode.Text;
                Screener.Name = txtName.Text;
                Screener.Exchange = txtExchange.Text;
                Screener.Limit = (int)numLimit.Value;

                SetFilteres();
                if (Screener.Filters.Count == 0) return;

                SetSignals();
                SetSort();

                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Screener error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }

        private void SetFilteres()
        {
            Screener.Filters.Clear();

            foreach (DataGridViewRow row in dataGridViewFilters.Rows)
            {
                Filter newFilter = new Filter();
                if (row.Cells[colField.Index].Value == null) continue;
                if (row.Cells[colOperation.Index].Value == null) throw new Exception("Select operation type from the list");
                if (row.Cells[colValue.Index].Value == null) throw new Exception("Select a value for the operation");

                switch (row.Cells[colField.Index].Value.ToString())
                {
                    case "market capitalization":
                        newFilter.Field = Field.MarketCapitalization;
                        break;
                    case "earnings share":
                        newFilter.Field = Field.EarningsShare;
                        break;
                    case "dividend yield":
                        newFilter.Field = Field.DividendYield;
                        break;
                    case "refund 1d p":
                        newFilter.Field = Field.Refund1dP;
                        break;
                    case "refund 5d p":
                        newFilter.Field = Field.Refund5dP;
                        break;
                    case "avgvol 1d":
                        newFilter.Field = Field.Refund5dP;
                        break;
                    case "avgvol 200d":
                        newFilter.Field = Field.Refund5dP;
                        break;

                    default:
                        throw new Exception("Select a field");
                }

                switch (row.Cells[colOperation.Index].Value.ToString())
                {
                    case "=":
                        newFilter.Operation = Operation.Equals;
                        break;
                    case ">":
                        newFilter.Operation = Operation.More;
                        break;
                    case "<":
                        newFilter.Operation = Operation.Less;
                        break;
                    case ">=":
                        newFilter.Operation = Operation.NotLess;
                        break;
                    case "<=":
                        newFilter.Operation = Operation.NotMore;
                        break;
                    case "!=":
                        newFilter.Operation = Operation.NotEquals;
                        break;

                    default:
                        throw new Exception("Select a operation");
                }

                newFilter.Value = row.Cells[colValue.Index].Value.ToString();
                Screener.Filters.Add(newFilter);
            }

            if (!string.IsNullOrEmpty(txtCode.Text))
            {
                Filter newFilter = new Filter()
                {
                    Field = Field.Code,
                    Operation = Operation.Equals,
                    Value = txtCode.Text,
                };
                Screener.Filters.Add(newFilter);
            }
            if (!string.IsNullOrEmpty(txtName.Text))
            {
                Filter newFilter = new Filter()
                {
                    Field = Field.Name,
                    Operation = Operation.Equals,
                    Value = txtName.Text,
                };
                Screener.Filters.Add(newFilter);
            }
            if (!string.IsNullOrEmpty(txtExchange.Text))
            {
                Filter newFilter = new Filter()
                {
                    Field = Field.Exchange,
                    Operation = Operation.Equals,
                    Value = txtExchange.Text,
                };
                Screener.Filters.Add(newFilter);
            }
            if (!string.IsNullOrEmpty(cboSector.Text))
            {
                Filter newFilter = new Filter()
                {
                    Field = Field.Sector,
                    Operation = Operation.Equals,
                    Value = cboSector.Text,
                };
                Screener.Filters.Add(newFilter);
            }
            if (!string.IsNullOrEmpty(cboIndustry.Text))
            {
                Filter newFilter = new Filter()
                {
                    Field = Field.Industry,
                    Operation = Operation.Equals,
                    Value = cboIndustry.Text,
                };
                Screener.Filters.Add(newFilter);
            }
            if (Screener.Filters.Count == 0)
            {
                MessageBox.Show("Not enough filters", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void SetSignals()
        {
            Screener.Signals = new List<Signal>();

            if (chk200d_new_lo.Checked) Screener.Signals.Add(Signal.New_200d_low);
            if (chk200d_new_hi.Checked) Screener.Signals.Add(Signal.New_200d_hi);
            if (chkBookvalue_neg.Checked) Screener.Signals.Add(Signal.Bookvalue_neg);
            if (chkBookvalue_pos.Checked) Screener.Signals.Add(Signal.Bookvalue_pos);
            if (chkWallstreet_lo.Checked) Screener.Signals.Add(Signal.Wallstreet_low);
            if (chkWallstreet_hi.Checked) Screener.Signals.Add(Signal.Wallstreet_hi);
            if (Screener.Signals.Count == 0) Screener.Signals = null;
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
                case "exchange":
                    field = Field.Exchange;
                    break;
                case "name":
                    field = Field.Name;
                    break;
                case "refund 1d":
                    field = Field.Refund1dP;
                    break;
                case "market capitalization":
                    field = Field.MarketCapitalization;
                    break;
                case "earnings share":
                    field = Field.EarningsShare;
                    break;
                case "dividend yield":
                    field = Field.DividendYield;
                    break;
                case "sector":
                    field = Field.Sector;
                    break;
                case "industry":
                    field = Field.Industry;
                    break;
                default:
                    Screener.Sort = null;
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

            Screener.Sort = (field, order);
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
                if (item.Item2 == cboSector.Text)
                {
                    list.Add(item.Item1);
                }
            }
            cboIndustry.Items.AddRange(list.Distinct().ToArray());
        }

        private void btnClearFilters_Click(object sender, EventArgs e)
        {
            cboSector.Text = null;
            cboIndustry.Text = null;
            txtCode.Text = null;
            txtName.Text = null;
            txtExchange.Text = null;
            numLimit.Value = 100;
            chk200d_new_hi.CheckState = CheckState.Unchecked;
            chk200d_new_lo.CheckState = CheckState.Unchecked;
            chkWallstreet_hi.CheckState = CheckState.Unchecked;
            chkWallstreet_lo.CheckState = CheckState.Unchecked;
            rbtnSortAsc.Checked = true;
            rbtnSortDesc.Checked = false;
            dataGridViewFilters.Rows.Clear();
        }

        private void ClearToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridViewFilters.Rows.RemoveAt(dataGridViewFilters.CurrentCell.RowIndex);
        }

        private void clearToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
    }
}
