import plotly.express as px
from dash import Dash, html, Input, Output, dcc, dash_table, callback
# import plotly.graph_objects as go
import pandas as pd
from sqlalchemy import create_engine
# from sqlalchemy.engine import URL
# import pyodbc

# Creating the connection string
Driver = 'SQL Server'
Server_Name = 'LAPTOP-HHQOVSI1\SQLEXPRESS'
Database = 'AdventureWorks2016'

connection_string = f'mssql+pyodbc://@{Server_Name}/{Database}?driver={Driver}'

# Creating the engine
engine = create_engine(connection_string)

# The SQL Queries Section

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
GeographicalMapQuery = '''
        SELECT cr.Name Country, cr.CountryRegionCode, count(soh.SalesOrderID) SalesOrdered, count(soh.CustomerID) Customers, count(soh.SalesPersonID)SalesPerson
        FROM Sales.SalesOrderHeader soh
        JOIN Person.Address a ON a.AddressID = soh.ShipToAddressID OR a.AddressID = soh.BillToAddressID
        JOIN Person.StateProvince sp ON sp.StateProvinceID = a.StateProvinceID
        JOIN Person.CountryRegion cr ON cr.CountryRegionCode = sp.CountryRegionCode
        GROUP BY cr.Name, cr.CountryRegionCode
        ORDER BY cr.Name;
        '''

SaleReasonQuery = '''
        SELECT sr.Name, count(sr.Name) Total
        FROM Sales.SalesOrderHeaderSalesReason hsr
        JOIN Sales.SalesReason sr ON sr.SalesReasonID = hsr.SalesReasonID
        GROUP BY sr.Name;
        '''
OnlinePhysicalQuery = '''
        SELECT soh.OnlineOrderFlag, count(soh.OnlineOrderFlag) Total 
        FROM Sales.SalesOrderHeader soh
        GROUP BY soh.OnlineOrderFlag;
        '''
CustomersByCountryQuery = '''
        SELECT cr.Name, count(soh.CustomerID) Customers
        FROM Sales.Customer c
        JOIN Person.Person p ON p.BusinessEntityID = c.PersonID
        JOIN Person.BusinessEntity be ON be.BusinessEntityID = p.BusinessEntityID
        JOIN Person.BusinessEntityAddress bea ON bea.BusinessEntityID = be.BusinessEntityID
        JOIN Person.Address a ON a.AddressID = bea.AddressID
        JOIN Person.StateProvince sp ON sp.StateProvinceID = a.StateProvinceID
        JOIN Person.CountryRegion cr On cr.CountryRegionCode = sp.CountryRegionCode
        JOIN Sales.SalesOrderHeader soh ON soh.BillToAddressID = a.AddressID
        GROUP BY cr.Name
        '''
CustomersByProvinceQuery = '''
        SELECT sp.Name, count(soh.CustomerID) Customers
        FROM Sales.Customer c
        JOIN Person.Person p ON p.BusinessEntityID = c.PersonID
        JOIN Person.BusinessEntity be ON be.BusinessEntityID = p.BusinessEntityID
        JOIN Person.BusinessEntityAddress bea ON bea.BusinessEntityID = be.BusinessEntityID
        JOIN Person.Address a ON a.AddressID = bea.AddressID
        JOIN Person.StateProvince sp ON sp.StateProvinceID = a.StateProvinceID
        JOIN Sales.SalesOrderHeader soh ON soh.BillToAddressID = a.AddressID
        GROUP BY sp.Name
        '''
CustomersByCityQuery = '''
        SELECT a.City, count(soh.CustomerID) Customers
        FROM Sales.Customer c
        JOIN Person.Person p ON p.BusinessEntityID = c.PersonID
        JOIN Person.BusinessEntity be ON be.BusinessEntityID = p.BusinessEntityID
        JOIN Person.BusinessEntityAddress bea ON bea.BusinessEntityID = be.BusinessEntityID
        JOIN Person.Address a ON a.AddressID = bea.AddressID
        JOIN Sales.SalesOrderHeader soh ON soh.BillToAddressID = a.AddressID
        GROUP BY a.City
        '''
SpecialOfferQuery = '''
        SELECT so.Type, count(so.Type) NumberOfSales
        FROM Sales.SalesOrderDetail sod
        JOIN Sales.SpecialOfferProduct sop ON sop.ProductID = sod.ProductID
        JOIN Sales.SpecialOffer so ON so.SpecialOfferID = sop.SpecialOfferID
        GROUP BY so.Type;
        '''
ProductMakeQuery = '''
        SELECT 
        CASE WHEN MakeFlag = 0 THEN 'Purchased' ELSE 'Manufactured' END Make, 
        count(MakeFlag) Total
        FROM Production.Product
        GROUP BY 
        CASE WHEN MakeFlag = 0 THEN 'Purchased' ELSE 'Manufactured' END;
        '''
ProductFinishedGoodsQuery = '''
        SELECT 
        CASE WHEN FinishedGoodsFlag = 0 THEN 'Not Salable' ELSE 'Salable' END FinishedGoods, 
        count(FinishedGoodsFlag) Total
        FROM Production.Product
        GROUP BY 
        CASE WHEN FinishedGoodsFlag = 0 THEN 'Not Salable' ELSE 'Salable' END;
        '''
ProductStyleQuery = '''
        SELECT 
        CASE WHEN Style='M' THEN 'MALE'
        WHEN Style= 'W' THEN 'FEMALE'
        WHEN Style='U' THEN 'UNION' 
        ELSE 'N/A' END Style, 
        count(*) Total
        FROM Production.Product p
        GROUP BY 
        CASE WHEN Style='M' THEN 'MALE'  
        WHEN Style= 'W' THEN 'FEMALE' 
        WHEN Style='U' THEN 'UNION'
        ELSE 'N/A' END;
        '''
ProductTransactionTypeQuery = '''
        SELECT 
        CASE WHEN transactiontype = 'W' THEN 'WorkOrder'
        WHEN TransactionType = 'S' THEN 'SalesOrder' 
        ELSE 'PurchaseOrder' END TransactionType,
        sum(quantity) Quantity, sum(actualcost) ActualCost, count(t.TransactionType) TransactionTypeTotal
        FROM Production.Product p
        Join Production.TransactionHistory t ON p.ProductID = t.ProductID
        GROUP BY 
        CASE WHEN transactiontype = 'W' THEN 'WorkOrder'
        WHEN TransactionType = 'S' THEN 'SalesOrder' 
        ELSE 'PurchaseOrder' END;
        '''
ProductClassQuery = '''
        SELECT 
        CASE WHEN Class='H' THEN 'High'
        WHEN Class='L' THEN 'Low' ELSE 'Medium' End Class,
        Count(*) TypesOfColors 
        FROM Production.Product 
        group by 
        CASE WHEN Class='H' THEN 'High'
        WHEN Class='L' THEN 'Low' ELSE 'Medium' End
        ORDER BY count(*)
        '''
# End SQL Queries


# The Dataframes section
connection = engine.connect()

TopSelectionSalesPersonDF = pd.read_sql_query(TopSelectionSalesPersonQuery, connection)
TopSelectionVendorDF = pd.read_sql_query(TopSelectionVendorQuery, connection)
TopSelectionCustomerDF = pd.read_sql_query(TopSelectionCustomerQuery, connection)
TopSelectionStoreDF = pd.read_sql_query(TopSelectionStoreQuery, connection)
TopProductDF = pd.read_sql_query(TopProductQuery, connection)
GeographicalMapDF = pd.read_sql_query(GeographicalMapQuery, connection)
SalesReasonDF = pd.read_sql_query(SaleReasonQuery, connection)
OnlinePhysicalDF = pd.read_sql_query(OnlinePhysicalQuery, connection)
CustomersByCityDF = pd.read_sql_query(CustomersByCityQuery, connection)
CustomersByCountryDF = pd.read_sql_query(CustomersByCountryQuery, connection)
CustomersByProvinceDF = pd.read_sql_query(CustomersByProvinceQuery, connection)
SpecialOfferDF = pd.read_sql_query(SpecialOfferQuery, connection)
ProductStyleDF = pd.read_sql_query(ProductStyleQuery, connection)
ProductMakeDF = pd.read_sql_query(ProductMakeQuery, connection)
ProductClassDF = pd.read_sql_query(ProductClassQuery, connection)
ProductFinishedGoodsDF = pd.read_sql_query(ProductFinishedGoodsQuery, connection)
ProductTransactionTypeDF = pd.read_sql_query(ProductTransactionTypeQuery, connection)

connection.close()
# Dataframe Section


# Initiating the application
external_stylesheet = ['style.css']
app = Dash(__name__, external_stylesheets=external_stylesheet)

app.layout = html.Div(className='parent1', children=[
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
            dcc.Graph(figure={}, id='topSelectionGraph')
        ])
    ]),
    html.Div(className='TopLow', children=[

    ]),
    html.Div(className='TopProduct', children=[
        html.H2('Top Product'),
        html.Div(className='TopProductHead', children=[
            dcc.Dropdown(options=['Top', 'Bottom'],
                         value='Top',
                         id='topProductDropdown1'),
            dcc.Dropdown(options=[10, 15, 20, 30, 50],
                         value=10,
                         id='topProductDropdown2'),
            html.P('Sold')
        ]),
        html.Div(id='TopProductGraphID', children=[
            # dcc.Graph(figure={}, id='topProductGraph')
            dash_table.DataTable(data=TopProductDF.to_dict('records'),
                                 page_size=15,
                                 id='topProductTable',
                                 columns=[{'id': c, 'name': c} for c in TopProductDF.columns],

                                 style_cell_conditional=[
                                    {
                                        'if': {'column_id': c},
                                        'textAlign': 'left'
                                    } for c in ['Name']
                                 ],
                                 style_data={
                                     'color': 'black',
                                     'background': 'white'
                                 },
                                 style_data_conditional=[
                                     {
                                         'if': {'row_index': 'odd'},
                                         'backgroundColor': 'rgb(220, 220, 220)',
                                     }
                                 ],
                                 style_header={
                                     'backgroundColor': 'rgb(210, 210, 210)',
                                     'color': 'black',
                                     'fontWeight': 'bold'
                                 }
                                 )
        ])
    ]),
    html.Div(className='VendorProduct', children=[

    ]),
    html.Div(className='GeographicalMap', children=[
        html.H2("Scatter-Geographic Charts"),
        html.Div([
            html.Div(className='GeographicHead', children=[
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
    html.Div(className='SaleReason', children=[
        html.H2('Sale Reason'),
        html.Div(
            dcc.Graph(figure={}, id='salesReasonGraph')
        )
    ]),
    html.Div(className='OnlinePhysical', children=[
        html.H2('Online vs Physical'),
        html.Div(
            dcc.Graph(figure={}, id='onlinePhysicalGraph')
        )
    ]),
    html.Div(className='Customer', children=[

    ]),
    html.Div(className='Sale', children=[

    ]),
    html.Div(className='SpecialOffer', children=[
        html.H2('Sales Special Offers'),
        html.Div(
            dcc.Graph(id='specialOfferGraph')
        )
    ]),
    html.Div(className='Product', children=[
        html.H2('Products'),
        html.Div(className='ProductMake', children=[
            html.H3('Product Make'),
            dcc.Graph(figure={}, id='productMakeGraph')
        ]),
        html.Div(className='ProductStyle', children=[
            html.H3("Product Style"),
            dcc.Graph(figure={}, id='productStyleGraph')
        ]),
        html.Div(className='ProductClass', children=[
            html.H3('Product Class'),
            dcc.Graph(figure={}, id='productClassGraph')
        ]),
        html.Div(className='ProductTransactionType', children=[
            html.H3('Product Transaction Type'),
            dcc.Graph(figure={}, id='productTransactionTypeGraph')
        ]),
        html.Div(className='ProductFinishedGoods', children=[
            html.H3('Product Finished Goods'),
            dcc.Graph(figure={}, id='productFinishedGoodsGraph')
        ])
    ]),
    html.Div(className='LoyalCustomer', children=[

    ])
], style={'marginBottom': 50, 'marginTop':25})

# End application initiation


# The callbacks for user interaction
#
#
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
#
#
# Top Product callback
@callback(
    Output(component_id='topProductTable', component_property='data'),
    Input(component_id='topProductDropdown1', component_property='value'),
    Input(component_id='topProductDropdown2', component_property='value'),
)
def top_product(value1, value2):
    if value1 == 'Bottom':
        query = f"""
        SELECT TOP {value2} p.Name , count(soh.SalesOrderID) SalesOrderID, sum(soh.TotalDue) TotalDue
        FROM Production.Product p
        JOIN Production.ProductProductPhoto ppp ON ppp.ProductID = p.ProductID
        JOIN Sales.SpecialOfferProduct sop ON sop.ProductID = ppp.ProductID
        JOIN Sales.SalesOrderDetail sod ON sod.SpecialOfferID = sop.SpecialOfferID
        JOIN Sales.SalesOrderHeader soh ON soh.SalesOrderID = sod.SalesOrderID
        GROUP BY P.Name
        ORDER BY count(soh.SalesOrderID) ASC;
        """
    else:
        query = f"""
        SELECT TOP {value2} p.Name , count(soh.SalesOrderID) SalesOrderID, sum(soh.TotalDue) TotalDue
        FROM Production.Product p
        JOIN Production.ProductProductPhoto ppp ON ppp.ProductID = p.ProductID
        JOIN Sales.SpecialOfferProduct sop ON sop.ProductID = ppp.ProductID
        JOIN Sales.SalesOrderDetail sod ON sod.SpecialOfferID = sop.SpecialOfferID
        JOIN Sales.SalesOrderHeader soh ON soh.SalesOrderID = sod.SalesOrderID
        GROUP BY P.Name
        ORDER BY count(soh.SalesOrderID) DESC;
        """

    connection = engine.connect()
    data = pd.read_sql_query(query, connection)
    connection.close()

    return data.to_dict('records')
#
# End of Top Product
#
#
# The callback for Geographical Map
@callback(
    Output(component_id='scatter-geo', component_property='figure'),
    Input(component_id='geo_dropdown', component_property='value')
)
def update_scatter_geo(value):
    fig = px.scatter_geo(GeographicalMapDF, locations="Country",
                         locationmode="country names",
                         size=value,
                         color="Country")

    return fig
# The end of callback for Geographical Map
#
#
#
# The callback for Sales Reason
@callback(
    Output(component_id='salesReasonGraph', component_property='figure'),
    Input(component_id='salesReasonGraph', component_property='id')
)
def sales_reason(value):
    fig = px.pie(SalesReasonDF, names='Name', values='Total', color='Name')

    return fig
# End callback for Sales reason
#
#
#
# The callback for Online or Physical
@callback(
    Output(component_id='onlinePhysicalGraph', component_property='figure'),
    Input(component_id='onlinePhysicalGraph', component_property='id')
)
def online_physical(value):
    fig = px.bar(OnlinePhysicalDF, x='OnlineOrderFlag', y='Total', color='OnlineOrderFlag')

    return fig
# End of callback for OnlinePhysical
#
#
#
# The callback for Special Offers
@callback(
    Output(component_id='specialOfferGraph', component_property='figure'),
    Input(component_id='specialOfferGraph', component_property='id')
)
def special_offer(value):
    fig = px.pie(SpecialOfferDF, values='NumberOfSales', names='Type', hole=0.6)

    return fig
# End call back for Special Offer
#
#
#
# Callback for Product
@callback(
    Output(component_id='productMakeGraph', component_property='figure'),
    Output(component_id='productStyleGraph', component_property='figure'),
    Output(component_id='productClassGraph', component_property='figure'),
    Output(component_id='productFinishedGoodsGraph', component_property='figure'),
    Output(component_id='productTransactionTypeGraph', component_property='figure'),
    Input(component_id='productMakeGraph', component_property='id')
)
def product(value):
    fig1 = px.histogram(ProductMakeDF, x='Make', y='Total')
    fig2 = px.pie(ProductStyleDF, names='Style', values='Total')
    fig3 = px.bar(ProductClassDF, x='Class', y='TypesOfColors')
    fig4 = px.bar(ProductFinishedGoodsDF, x='FinishedGoods', y='Total')
    fig5 = px.pie(ProductTransactionTypeDF, values='TransactionTypeTotal', names='TransactionType')

    return fig1, fig2, fig3, fig4, fig5


# End of callbacks


# Calling the main/ Running the app

if __name__ == '__main__':
    app.run(debug=True)

