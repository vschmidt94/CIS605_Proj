'Template Layout Copyright (c) 2009-2015 Dan Turk
'All other code produced by Vaughan Schmidt.

#Region "Class / File Comment Header block"
'Program:               CIS605-Project-Schmidt-Vaughan      
'File:                  FrmMain.vb
'Author:                Vaughan Schmidt
'Description:           FrmMain.vb
'                       The main user interface form file for project.
'Date:                  12 Sept 2015
'                        - Initial file creation, initial layout
'Tier:                  User Interface  
'Exceptions:            None  
'Exception-Handling:    None 
'Events:                Basic UI system events 
'Event-Handling:        By UI event handlers
#End Region 'Class / File Comment Header block

#Region "Option / Imports"
Option Explicit On      'Must declare variables before using them
Option Strict On        'Must perform explicit data type conversions
#End Region 'Option / Imports

Public Class FrmMain

#Region "Attributes"
    '******************************************************************
    'Attributes + Module-level Constants+Variables
    '******************************************************************

    'No Attributes are currently defined.

    '********** Module-level constants

    '********** Module-level variables
    Private mThemePark As ThemePark

#End Region 'Attributes

#Region "Constructors"
    '******************************************************************
    'Constructors
    '******************************************************************

    'No Constructors are currently defined.
    'These are all public.

    '********** Default constructor
    '             - no parameters

    '********** Special constructor(s)
    '             - typically constructors have parameters 
    '               that are used to initialize attributes

    '********** Copy constructor(s)
    '             - one parameter, an object of the same class

#End Region 'Constructors

#Region "Get/Set Methods"
    '******************************************************************
    'Get/Set Methods
    '******************************************************************

    'No Get/Set Methods are currently defined.

    '********** Public Get/Set Methods
    '             - call private get/set methods to implement

    '********** Private Get/Set Methods
    '             - access attributes, begin name with underscore (_)

#End Region 'Get/Set Methods

#Region "Behavioral Methods"
    '******************************************************************
    'Behavioral Methods
    '******************************************************************

    'No Behavioral Methods are currently defined.

    '********** Public Shared Behavioral Methods

    '********** Private Shared Behavioral Methods
    ''' -------------------------------------------------------------
    ''' <summary>
    ''' Private shared exit function
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -------------------------------------------------------------
    Private Sub _exit(
            ByVal sender As Object,
            ByVal e As EventArgs
            )

        Me.Close()

    End Sub '_exit()

    '********** Public Non-Shared Behavioral Methods

    '********** Private Non-Shared Behavioral Methods

    Private Sub _initializeBusinessLogic()

        'Create a Theme Park object
        mThemePark = New ThemePark()

        'Log the initializtion

        txtTrxLogTabLog.Text &= vbCrLf & CType(TimeValue(CType(Now, String)), String) _
            & " - ThemePark Object initialized"


    End Sub '_initializeBusinessLogic()

    Private Sub _initializeUserInterface()

        Me.Text = "Cretaceous Park System Portal"

        'Set the cancel (ESC) button to click the btnExit
        Me.CancelButton = btnExit

        'Populate the dashboard
        txtNumCustomersTabDashboard.Text = CType(mThemePark.numCustomers(), String)
        txtNumPassbooksTabDashboard.Text = CType(mThemePark.numPassbooks(), String)
        txtNumFeaturesTabDashboard.Text = CType(mThemePark.numFeatures(), String)

        'Log the initialization
        txtTrxLogTabLog.Text &= vbCrLf & CType(TimeValue(CType(Now, String)), String) _
            & " - User Interface Intialized"
        txtTrxLogTabLog.Text &= vbCrLf & CType(TimeValue(CType(Now, String)), String) _
            & " - Dashboard initial population complete"

    End Sub '_initalizeUserInterface()

#End Region 'Behavioral Methods

#Region "Event Procedures"
    '******************************************************************
    'Event Procedures
    '******************************************************************

    'No Event Procedures are currently defined.
    'These are all private.

    '********** User-Interface Event Procedures
    '             - Initiated explicitly by user

    ''' <summary>
    ''' Changes the availability of tabs on main view based on the
    ''' user roll selected.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub cmbRoleSelector_SelectedIndexChanged(
            sender As Object,
            e As EventArgs
            ) _
        Handles cmbRoleSelector.SelectedIndexChanged

        'TODO: Highlight tab pages appropriate to user role /
        '      lock out pages not appropriate to user role.
        '      Not explicitly required but would like to try
        '      to do this.
    End Sub

    ''' <summary>
    ''' Exits the program.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnExit_Click(
            sender As Object,
            e As EventArgs
            ) _
        Handles btnExit.Click

        Me.Close()

    End Sub 'btnExit_Click()

    Private Sub btnAddCustomer_Click(
            sender As Object,
            e As EventArgs
            ) _
        Handles btnAddNewCustTabNewCustomer.Click

        Dim _customer As Customer
        Dim _custName As String
        Dim _custID As String

        'Parse the user's input from new customer tab
        'TODO: do these try/catch blocks ever catch anything?  Maybe add some other 
        'checks for robustness
        Try
            _custName = txtAddCustNameTabNewCustomer.Text
        Catch ex As Exception
            MessageBox.Show("Please enter a valid customer name.")
            txtAddCustNameTabNewCustomer.SelectAll()
            txtAddCustNameTabNewCustomer.Focus()
            Exit Sub
        End Try

        Try
            _custID = txtAddCustIDTabNewCustomer.Text
        Catch ex As Exception
            MessageBox.Show("Please enter a valid customer ID.")
            txtAddCustIDTabNewCustomer.SelectAll()
            txtAddCustIDTabNewCustomer.Focus()
            Exit Sub
        End Try

        'Create a new Customer
        _customer = New Customer(_custID, _custName)

        'Add 1 to the Customer count & update dashpboard
        mThemePark.numCustomers = mThemePark.numCustomers + 1
        txtNumCustomersTabDashboard.Text = CType(mThemePark.numCustomers, String)

        'Confirm with message box and clear the form for more
        MessageBox.Show("New Customer Added")
        txtAddCustIDTabNewCustomer.ResetText()
        txtAddCustNameTabNewCustomer.ResetText()

        'Add to the log
        txtTrxLogTabLog.Text &= vbCrLf & CType(TimeValue(CType(Now, String)), String) _
            & " - New Customer " & _custID & " added"


    End Sub 'btnAddCustomer_Click()

    Private Sub btnAddFeature_Click(
            sender As Object,
            e As EventArgs
            ) _
        Handles btnAddFeatureTabDefineFeature.Click

        Dim _feature As Feature
        Dim _featureName As String
        Dim _featureID As String
        Dim _featureUOM As String
        Dim _featureAdultPrice As Decimal
        Dim _featureChildPrice As Decimal

        'Parse and validate the form input.
        'For sake of time, not validating string entries right now, focus on stuff that
        'can be parsed.
        Try
            _featureAdultPrice = Decimal.Parse(txtFeatureAdultPriceTabDefineFeature.Text)
        Catch ex As Exception
            MessageBox.Show("Please enter a valid decimal price (1.23)")
            txtFeatureAdultPriceTabDefineFeature.SelectAll()
            txtFeatureAdultPriceTabDefineFeature.Focus()
            txtTrxLogTabLog.Text &= vbCrLf & "Invalid Feature Adult Price Entered"
            Exit Sub
        End Try

        Try
            _featureChildPrice = Decimal.Parse(txtFeatureChildPriceTabDefineFeature.Text)
        Catch ex As Exception
            MessageBox.Show("Please enter a valid decimal price (1.23)")
            txtFeatureChildPriceTabDefineFeature.SelectAll()
            txtFeatureChildPriceTabDefineFeature.Focus()
            txtTrxLogTabLog.Text &= vbCrLf & "Invalid Feature Child Price Entered"
            Exit Sub
        End Try

        'TODO:  Add some text validation for strings...
        _featureName = txtNewFeatureNameTabDefineFeature.Text
        _featureUOM = txtFeatureUOMTabDefineFeature.Text
        _featureID = txtNewFeatureIDTabDefineFeature.Text

        'Create New Feature
        _feature = New Feature(_featureID,
                               _featureName,
                               _featureUOM,
                               _featureAdultPrice,
                               _featureChildPrice)

        'Add 1 to the Feature Count & update dashboard
        mThemePark.numFeatures = mThemePark.numFeatures + 1
        txtNumFeaturesTabDashboard.Text = CType(mThemePark.numFeatures, String)

        'Confirm with message box and clear the form for more
        MessageBox.Show("New Feature Added")
        txtNewFeatureNameTabDefineFeature.ResetText()
        txtNewFeatureIDTabDefineFeature.ResetText()
        txtFeatureUOMTabDefineFeature.ResetText()
        txtFeatureAdultPriceTabDefineFeature.ResetText()
        txtFeatureChildPriceTabDefineFeature.ResetText()

        'Add to the log
        txtTrxLogTabLog.Text &= vbCrLf & CType(TimeValue(CType(Now, String)), String) _
            & " - New Feature " & _featureID & " added"
    End Sub 'btnAddFeature_Click() 

    ''' <summary> 
    ''' Populates some test data 
    ''' </summary> 
    ''' <param name="sender"></param> 
    ''' <param name="e"></param> 
    Private Sub btnAddTestData_Click(sender As Object, e As EventArgs) _
        Handles btnProcessTestData.Click

        Dim themePark As ThemePark
        Dim customerC01 As Customer
        Dim customerC02 As Customer
        Dim featureF01 As Feature
        Dim featureF02 As Feature
        Dim featureF03 As Feature
        Dim passbook As Passbook
        Dim passbookFeature As PassbookFeature
        Dim usedFeature As UsedFeature

        'Log a message that test data is being loaded
        txtTrxLogTabLog.Text &= vbCrLf & CType(TimeValue(CType(Now, String)), String) _
            & " - LOADING TEST DATA STARTED"

        'Create Theme Park, ID "CIS605 Theme Park"
        themePark = New ThemePark("CIS605 Theme Park")
        txtTrxLogTabLog.Text &= vbCrLf & CType(TimeValue(CType(Now, String)), String) _
            & "New Theme Park Added: " _
            & themePark.ToString

        'Create feature, ID “F01”, Description “Park Pass”, Units “Day”, Adult Price $100, Child Price $80
        featureF01 = New Feature("F01", "Park Pass", "Day", 100, 80)
        txtTrxLogTabLog.Text &= vbCrLf & CType(TimeValue(CType(Now, String)), String) _
            & "New Feature Added: " _
            & featureF01.ToString

        'Create Feature, ID “F02”, Description “Early Entry Pass”, Units “Day”, Adult Price $10, Child Price $5
        featureF02 = New Feature("F02", "Early Entry Pass", "Day", 10, 5)
        txtTrxLogTabLog.Text &= vbCrLf & CType(TimeValue(CType(Now, String)), String) _
            & "New Feature Added: " _
            & featureF02.ToString

        'Create Feature, ID “F03”, Description “Meal Plan”, Units “Meal”, Adult Price $30, Child Price $20
        featureF03 = New Feature("F02", "Meal Plan", "Meal", 30, 20)
        txtTrxLogTabLog.Text &= vbCrLf & CType(TimeValue(CType(Now, String)), String) _
            & "New Feature Added: " _
            & featureF03.ToString

        'Create Customer, ID “C01”, Name “CName01”
        customerC01 = New Customer("C01", "CName01")
        txtTrxLogTabLog.Text &= vbCrLf & CType(TimeValue(CType(Now, String)), String) _
            & "New Feature Added: " _
            & featureF03.ToString
        'Create Customer, ID “C02”, Name “CName02”
        'Create Customer, ID “C03”, Name “Customer Name 03”
        'Create Passbook, ID “PB01”, Customer “C01” reference, DatePurch 9/15/2015, Visitor Name “self”, Visitor BDay 1/1/1980
        'Create Passbook, ID “PB02”, Customer “C02” reference, DatePurch 9/16/2015, Visitor Name “self”, Visitor BDay 6/1/1985
        'Create Passbook, ID “PB03”, Customer “C02” reference, DatePurch 9/17/2015, Visitor Name “C02 Visitor”, Visitor BDay 12/1/2003
        'Create Passbook, ID “PB04”, Customer “C03” reference, DatePurch 8/15/2015, Visitor Name “self”, Visitor BDay 1/1/1975
        'Create Passbook, ID “PB05”, Customer “C03” reference, DatePurch 9/15/2015, Visitor Name “C03 Visitor 1”, Visitor BDay 10/7/2002
        'Create Passbook, ID “PB06”, Customer “C03” reference, DatePurch 10/15/2015, Visitor Name “C03 Visitor 2”, Visitor BDay 10/8/2002
        'Purchase Feature (i.e. Create PassbookFeature), ID “PBF01”, Qty 1, Passbook “PB01” reference, feature “F01” reference
        'Purchase Feature (i.e. Create PassbookFeature), ID “PBF02”, Qty 2, Passbook “PB02” reference, feature “F01” reference
        'Purchase Feature (i.e. Create PassbookFeature), ID “PBF03”, Qty 3, Passbook “PB03” reference, feature “F01” reference
        'Purchase Feature (i.e. Create PassbookFeature), ID “PBF04”, Qty 1, Passbook “PB04” reference, feature “F01” reference
        'Purchase Feature (i.e. Create PassbookFeature), ID “PBF05”, Qty 1, Passbook “PB05” reference, feature “F01” reference
        'Purchase Feature (i.e. Create PassbookFeature), ID “PBF06”, Qty 1, Passbook “PB06” reference, feature “F01” reference
        'Purchase Feature (i.e. Create PassbookFeature), ID “PBF07”, Qty 3, Passbook “PB03” reference, feature “F02” reference
        'Purchase Feature (i.e. Create PassbookFeature), ID “PBF08”, Qty 9, Passbook “PB03” reference, feature “F03” reference
        'Purchase Feature (i.e. Create PassbookFeature), ID “PBF09”, Qty 1, Passbook “PB04” reference, feature “F01” reference
        'Purchase Feature (i.e. Create PassbookFeature), ID “PBF10”, Qty 3, Passbook “PB04” reference, feature “F01” reference
        'Use Feature (i.e. Create UsedFeature), ID “UF01”, PBFeature “PBF01” reference, DateUsed 10/20/2015, LocationUsed “Epcot Center”, QtyUsed 1

    End Sub

    '********** User-Interface Event Procedures
    '             - Initiated automatically by system

    ''' -------------------------------------------------------------
    ''' <summary>
    ''' _FrmMain_Load() will initialize the Business Logic and the
    ''' User Interface bu calling the desired initialization
    ''' functions.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -------------------------------------------------------------
    Private Sub _FrmMain_Load(
            ByVal sender As Object,
            ByVal e As EventArgs
            ) _
        Handles MyBase.Load

        _initializeBusinessLogic()
        _initializeUserInterface()

    End Sub '_FrmMain_Load()

    Private Sub _tab_Index_Changed(
            sender As Object,
            e As EventArgs
            ) _
           Handles tbcMainActivities.SelectedIndexChanged

    End Sub '_tab_Index_Changed()

    Private Sub chkIsChildTabBuyFeature_CheckedChanged(sender As Object, e As EventArgs) Handles chkIsChildTabBuyFeature.CheckedChanged

    End Sub



    '********** Business Logic Event Procedures
    '             - Initiated as a result of business logic
    '               method(s) running

#End Region 'Event Procedures

#Region "Events"
    '******************************************************************
    'Events
    '******************************************************************

    'No Events are currently defined.
    'These are all public.

#End Region 'Events 

End Class 'FrmMain
