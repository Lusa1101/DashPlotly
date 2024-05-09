import plotly.express as px
from dash import Dash, html, Input, Output, dcc, callback
import pandas as pd
from sqlalchemy import create_engine


# connection = pyodbc.connect('DRIVER={SQL Server};'+
#                             'Server=LAPTOP-HHQOVSI1\SQLEXPRESS;'+
#                             'Database=AdventureWorks2019;'+
#                             'Trusted_Connection=True')
# print("Connected to db.")

SERVERNAME = r'LAPTOP-HHQOVSI1\SQLEXPRESS'
DRIVER = 'SQL Server'
DATABASE = 'AdventureWorks2016'

connection_string = f'mssql+pyodbc://@{SERVERNAME}/{DATABASE}?driver={DRIVER}'
engine = create_engine(connection_string)

# Arrays for the table under/in each schema defined in the application
Sales = ['Store', 'Customer', 'SalesOrderDetail', 'SalesOrderHeader', 'Currency', 'CurrencyRate', 'SalesPerson',
         'SalesReason', 'SpecialOffer', 'SalesTerritory', 'SalesTerritoryHistory']
Purchasing = ['ProductVendor', 'Vendor', 'ShipMethod', 'PurchaseOrderDetail', 'PurchaseOrderHeader']
Production = ['Location', 'ProductReview', 'Product', 'ProductModel', 'TransactionHistory', 'WorkOrder',
              'ProductDocument', 'ProductInventory', 'ProductCategory', 'ProductListPriceHistory',
              'WorkOrdeRouting', 'ProductDescription', 'ProductSubCategory']
Person = ['Address', 'AddressType', 'BusinessEntity', 'BusinessEntityType', 'BusinessEntityContact', 'ContactType',
          'CountryRegion', 'EmailAddress', 'Password', 'Person', 'PersonPhone', 'PhoneNumberType', 'StateProvince']
HumanResources = ['Department', 'Employee', 'EmployeeDepartmentHistory', 'EmployeePayHistory', 'JobCandidate', 'Shift']
# All the functions to be used in the application


def return_options(value):
    if value == 'Sales':
        return Sales
    elif value == 'Purchasing':
        return Purchasing
    elif value == 'Person':
        return Person
    elif value == 'HumanResources':
        return HumanResources
    else:
        return Production


# The queries to be used in this app
query1 = f'''
SELECT ppc.Name ,
round(SUM(sod.OrderQty * sod.UnitPrice ),2) AS TotalPurchase
FROM sales.SalesOrderDetail sod
JOIN production.product pp ON sod.ProductID = pp.ProductID
JOIN production.ProductSubcategory ppsc 
ON ppsc.ProductSubcategoryID = pp.ProductSubcategoryID
JOIN production.ProductCategory ppc 
ON ppc.ProductCategoryID = ppsc.ProductCategoryID
GROUP BY ppc.name
Order By SUM(sod.OrderQty * sod.UnitPrice) DESC;
'''

query2 = '''
SELECT st.Name, sth.StartDate, 
CASE WHEN sth.EndDate=NULL THEN FORMAT (GETDATE(), 'dd/MM/yyyy') ELSE FORMAT (sth.EndDate, 'dd MM yyyy') END EndDate
FROM Sales.SalesTerritoryHistory sth
JOIN Sales.SalesTerritory st ON sth.TerritoryID = st.TerritoryID
JOIN Sales.SalesPerson sp ON st.TerritoryID = sp.TerritoryID'''
#

query3 = '''
SELECT a.City, sum(TotalDue * AverageRate) Total
FROM Sales.SalesOrderHeader soh
JOIN Sales.Customer c ON c.CustomerID = soh.CustomerID
JOIN Sales.Store s ON s.BusinessEntityID = c.StoreID
JOIN Person.BusinessEntity be ON s.BusinessEntityID = be.BusinessEntityID
JOIN Person.BusinessEntityAddress bea ON bea.BusinessEntityID = be.BusinessEntityID
JOIN Person.Address a ON a.AddressID = bea.AddressID
JOIN Person.StateProvince sp ON sp.StateProvinceID = a.StateProvinceID
JOIN Person.CountryRegion cr ON cr.CountryRegionCode = sp.CountryRegionCode
JOIN Sales.CurrencyRate crt ON crt.CurrencyRateID = soh.CurrencyRateID
GROUP BY a.City
'''

query4 = '''
SELECT cr.Name Country, cr.CountryRegionCode, count(soh.SalesOrderID) SalesOrdered, count(soh.CustomerID) Customers, count(soh.SalesPersonID)SalesPerson
FROM Sales.SalesOrderHeader soh
JOIN Person.Address a ON a.AddressID = soh.ShipToAddressID OR a.AddressID = soh.BillToAddressID
JOIN Person.StateProvince sp ON sp.StateProvinceID = a.StateProvinceID
JOIN Person.CountryRegion cr ON cr.CountryRegionCode = sp.CountryRegionCode
GROUP BY cr.Name, cr.CountryRegionCode
ORDER BY cr.Name;
'''

salesDistributionQuery = '''
    SELECT ppsc.Name ,
round(SUM(sod.OrderQty * sod.UnitPrice ),2) AS TotalPurchase
FROM sales.SalesOrderDetail sod
JOIN production.product pp ON sod.ProductID = pp.ProductID
JOIN production.ProductSubcategory ppsc 
ON ppsc.ProductSubcategoryID = pp.ProductSubcategoryID
GROUP BY ppsc.name
ORDER BY SUM(sod.OrderQty * sod.UnitPrice) DESC;
'''

SpecialOfferQuery = '''
        SELECT so.Type, count(so.Type) NumberOfSales
        FROM Sales.SalesOrderDetail sod
        JOIN Sales.SpecialOfferProduct sop ON sop.ProductID = sod.ProductID
        JOIN Sales.SpecialOffer so ON so.SpecialOfferID = sop.SpecialOfferID
        GROUP BY so.Type;
        '''

TopSelectionVendorQuery = '''
    SELECT TOP 10 v.name VendorName, sum(poh.TotalDue) TotalPrice
    FROM Purchasing.PurchaseOrderHeader poh
    JOIN Purchasing.Vendor v ON poh.VendorID = v.BusinessEntityID
    GROUP BY v.Name
    ORDER BY TotalPrice DESC;
    '''
TopSelectionSalesPersonQuery = '''
    SELECT TOP 10 p.FirstName + ' ' + p.LastName SalesPersonNames, sum(soh.TotalDue) TotalDue
    FROM Sales.SalesOrderHeader soh
    JOIN Sales.SalesPerson sp ON soh.SalesPersonID = sp.BusinessEntityID
    JOIN HumanResources.Employee e ON e.BusinessEntityID = sp.BusinessEntityID
    JOIN Person.Person p ON p.BusinessEntityID = e.BusinessEntityID
    GROUP BY p.FirstName + ' ' + p.LastName
    ORDER BY sum(soh.TotalDue) DESC;
    '''
TopSelectionStoreQuery = '''
    SELECT TOP 10 s.Name StoreName, sum(soh.TotalDue) TotalDue
    FROM Sales.Store s
    JOIN Sales.Customer c ON c.StoreID = s.BusinessEntityID
    JOIN Sales.SalesOrderHeader soh ON soh.CustomerID = c.CustomerID
    GROUP BY s.Name
    ORDER BY sum(soh.TotalDue) DESC;
    '''
TopSelectionCustomerQuery = '''
    SELECT TOP 10 p.FirstName + ' ' + p.LastName CustomerNames, sum(soh.TotalDue) TotalDue
    FROM Sales.Customer c
    JOIN Sales.SalesOrderHeader soh ON soh.CustomerID = c.CustomerID
    JOIN Person.Person p ON p.BusinessEntityID = c.CustomerID
    GROUP BY p.FirstName + ' ' + p.LastName
    ORDER BY sum(soh.TotalDue) DESC;
    '''
TopProductQuery = f"""
        SELECT TOP 10 p.Name , count(soh.SalesOrderID) SalesOrderID, sum(soh.TotalDue) TotalDue
        FROM Production.Product p
        JOIN Production.ProductProductPhoto ppp ON ppp.ProductID = p.ProductID
        JOIN Sales.SpecialOfferProduct sop ON sop.ProductID = ppp.ProductID
        JOIN Sales.SalesOrderDetail sod ON sod.SpecialOfferID = sop.SpecialOfferID
        JOIN Sales.SalesOrderHeader soh ON soh.SalesOrderID = sod.SalesOrderID
        GROUP BY P.Name
        ORDER BY count(soh.SalesOrderID) DESC;
        """
SaleReasonQuery = '''
        SELECT sr.Name, count(sr.Name) Total
        FROM Sales.SalesOrderHeaderSalesReason hsr
        JOIN Sales.SalesReason sr ON sr.SalesReasonID = hsr.SalesReasonID
        GROUP BY sr.Name;
        '''
line_graph_query = f'''
      SELECT YEAR(soh.OrderDate) AS Year,
      CASE WHEN MONTH(soh.OrderDate) = 1 THEN 'JAN'
      WHEN MONTH(soh.OrderDate) = 2 THEN 'FEB'
      WHEN MONTH(soh.OrderDate) = 3 THEN 'MAR'
      WHEN MONTH(soh.OrderDate) = 4 THEN 'APR'
      WHEN MONTH(soh.OrderDate) = 5 THEN 'MAY'
      WHEN MONTH(soh.OrderDate) = 6 THEN 'JUN'
      WHEN MONTH(soh.OrderDate) = 7 THEN 'JUL'
      WHEN MONTH(soh.OrderDate) = 8 THEN 'AUG'
      WHEN MONTH(soh.OrderDate) = 9 THEN 'SEP'
      WHEN MONTH(soh.OrderDate) = 10 THEN 'OCT'
      WHEN MONTH(soh.OrderDate) = 11 THEN 'NOV'
      ELSE 'DEC' END Month,
      ROUND(SUM(soh.TotalDue),2) AS TotalSales,
      ROUND(SUM(sod.orderqty * sod.unitprice),2) AS TotalPurchases 
      FROM sales.SalesOrderDetail sod
      JOIN sales.SalesOrderHeader soh ON sod.SalesOrderID = soh.SalesOrderID
      GROUP BY YEAR(soh.OrderDate), MONTH(soh.OrderDate)
      ORDER BY MONTH(soh.OrderDate);
  '''

# Creating the dataframes
# Retrieve data from the sql server
connection = engine.connect()
gantt_df = pd.read_sql_query(query2, connection)
pie_df = pd.read_sql_query(query1, connection)
geo_df = pd.read_sql_query(query4, connection)
table_df = pd.read_sql_query(query4, connection)
TopSelectionSalesPersonDF = pd.read_sql_query(TopSelectionSalesPersonQuery, connection)
TopSelectionVendorDF = pd.read_sql_query(TopSelectionVendorQuery, connection)
TopSelectionCustomerDF = pd.read_sql_query(TopSelectionCustomerQuery, connection)
TopSelectionStoreDF = pd.read_sql_query(TopSelectionStoreQuery, connection)
TopProductDF = pd.read_sql_query(TopProductQuery, connection)
SalesReasonDF = pd.read_sql_query(SaleReasonQuery, connection)
SpecialOfferDF = pd.read_sql_query(SpecialOfferQuery, connection)
line_df = pd.read_sql_query(line_graph_query, connection)

# line_df = pd.read_sql_query(lineGraphQuery, connection)
connection.close()

# Initializing the app
external_stylesheets = ['https://codepen.io/chriddyp/pen/bWLwgP.css']
external_stylesheet = [r'C:\Users\omphu\PycharmProjects\pythonProject\style.css']
app = Dash(__name__, external_stylesheets=external_stylesheet)

# Setting up the colours
colors = {
    'background': '#222A34',
    'text': '#7DFFE5'
}

# Create the app layout
app.layout = html.Div(className='Parent', children=[
    html.H1("AdventureWorks2016 Data Report"),
    html.Div(className='TopSelectionClass', children=[
        html.H2('Top Selection'),
        html.Div(className='TopSelectionHead', children=[
            html.P("Top 10"),
            dcc.Dropdown(options=['Vendor', 'Store', 'SalesPerson', 'Customer'],
                         value='Vendor',
                         id='topSelection')
        ]),
        html.Div(id='topSelectionGraphID', children=[
            dcc.Graph(figure={}, id='topSelectionGraph'),
        ])
    ]),
    html.Div(className='SaleReason', children=[
        html.H2('Sale Reason'),
        html.Div(
            dcc.Graph(figure={}, id='salesReasonGraph')
        )
    ]),
    html.Div(children=[
        html.Div([
            html.H3("Scatter-Geographic Charts"),
            html.Div(className='row', children=[
                dcc.RadioItems(
                    id='geo_dropdown',
                    options=['SalesOrdered', 'Customers', 'SalesPerson'],
                    value='SalesPerson')
            ])
        ]),
        html.Div(
            dcc.Graph(figure={}, id='scatter-geo'),
        )
    ]),
    html.Div(className='SpecialOffer', children=[
        html.H2('Sales Special Offers'),
        html.Div(
            dcc.Graph(id='specialOfferGraph')
        )
    ]),
    html.Div(children=[
        html.Div(
            html.H3("Purchases by category")
        ),
        html.Div(
            dcc.Graph(figure={}, id='pie-chart')
        )
    ]),
    html.Div(className='row', children=[
        html.H3("TOTAL SALES AMOUNT PER MONTH (EACH YEAR)"),
        html.Div(
            dcc.Dropdown(
                options=['TotalSales', 'TotalPurchases'],
                value='TotalPurchases',
                id='line-dropdown',
                clearable=False
            )
        )
    ]),
    html.Div(
        dcc.Graph(id='line-graph')
    )
])


# Add controls to build the interaction
# Callback for Top Selection
@callback(
    Output(component_id='topSelectionGraph', component_property='figure'),
    Input(component_id='topSelection', component_property='value')
)
def top_selection(value):
    if value == 'Vendor':
        data = TopSelectionVendorDF.to_dict()
        fig = px.bar(data, x='VendorName', y='TotalPrice')

    elif value == 'Store':
        data = TopSelectionStoreDF.to_dict()
        fig = px.bar(data, x='StoreName', y='TotalDue')

    elif value == 'SalesPerson':
        data = TopSelectionSalesPersonDF.to_dict()
        fig = px.bar(data, x='SalesPersonNames', y='TotalDue')

    else:
        data = TopSelectionCustomerDF.to_dict()
        fig = px.bar(data, x='CustomerNames', y='TotalDue')

    return fig
# End of Top Selection


# The callback for Sales Reason
@callback(
    Output(component_id='salesReasonGraph', component_property='figure'),
    Input(component_id='salesReasonGraph', component_property='id')
)
def sales_reason(value):
    fig = px.pie(SalesReasonDF, names='Name', values='Total', color='Name')

    return fig
# End callback for Sales reason


# The callback for Special Offers
@callback(
    Output(component_id='specialOfferGraph', component_property='figure'),
    Input(component_id='specialOfferGraph', component_property='id')
)
def special_offer(value):
    fig = px.pie(SpecialOfferDF, values='NumberOfSales', names='Type', hole=0.6)

    return fig
# End call back for Special Offer


# The callback for Geographical Map
@callback(
    Output(component_id='scatter-geo', component_property='figure'),
    Input(component_id='geo_dropdown', component_property='value')
)
def update_scatter_geo(value):
    fig = px.scatter_geo(geo_df, locations="Country",
                         locationmode="country names",
                         size=value,
                         color="Country")
    return fig
# End callback for Geographical Map


# Callback for line-graph
@callback(
    Output(component_id='line-graph', component_property='figure'),
    Input(component_id='line-dropdown', component_property='value')
)
def upgrade_line_graph(value):
    # df1 = line_df[line_df['Year'] == f'{value}']
    fig = px.bar(line_df, x='Month', y=value, color='Year',
                 barmode='group')

    return fig
# End callback for


@callback(
    Output(component_id='pie-chart', component_property='figure'),
    Input(component_id='pie-chart', component_property='id'),
)
def update_pie_chart(value):
    data2 = pie_df.to_dict()
    fig2 = px.pie(data2, values='TotalPurchase', names='Name', hole=0.6)

    return fig2


# Run application
if __name__ == '__main__':
    app.run(debug=True)


# See PyCharm help at https://www.jetbrains.com/help/pycharm/
