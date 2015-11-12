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
    Private WithEvents mThemePark As ThemePark

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
        txtNumPassbookFeatureTabDashboard.Text = CType(mThemePark.numPassbookFeatures(), String)
        txtNumUsedFeatureTabDashboard.Text = CType(mThemePark.numUsedFeatures(), String)

        'Clear the combo boxes
        cboPassbookOwnerTabPurchasePassbook.Items.Clear()
        cboPickPassbookIDTabBuyFeature.Items.Clear()
        cboFeatureSelectorTabPostUsedFeature.Items.Clear()
        cboFeatureSelectTabBuyFeature.Items.Clear()
        cboPassbookIDTabPostUsedFeature.Items.Clear()
        cboPassbookTabUpdatePassbook.Items.Clear()

        'Log the initialization
        txtTrxLogTabLog.Text &= vbCrLf & CType(TimeValue(CType(Now, String)), String) _
            & " - User Interface Intialized" _
            & vbCrLf
        txtTrxLogTabLog.Text &= vbCrLf & CType(TimeValue(CType(Now, String)), String) _
            & " - Dashboard initial population complete" _
            & vbCrLf

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
        Dim customerC03 As Customer
        Dim featureF01 As Feature
        Dim featureF02 As Feature
        Dim featureF03 As Feature
        Dim passbookPB01 As Passbook
        Dim passbookPB02 As Passbook
        Dim passbookPB03 As Passbook
        Dim passbookPB04 As Passbook
        Dim passbookPB05 As Passbook
        Dim passbookPB06 As Passbook
        Dim pbf01 As PassbookFeature
        Dim pbf02 As PassbookFeature
        Dim pbf03 As PassbookFeature
        Dim pbf04 As PassbookFeature
        Dim pbf05 As PassbookFeature
        Dim pbf06 As PassbookFeature
        Dim pbf07 As PassbookFeature
        Dim pbf08 As PassbookFeature
        Dim pbf09 As PassbookFeature
        Dim pbf10 As PassbookFeature
        Dim uf01 As UsedFeature
        Dim uf02 As UsedFeature
        Dim uf03 As UsedFeature
        Dim uf04 As UsedFeature

        'Log a message that test data is being loaded
        txtTrxLogTabLog.Text &= vbCrLf & CType(TimeValue(CType(Now, String)), String) _
            & " - LOADING TEST DATA STARTED" _
            & vbCrLf

        'Create Theme Park, ID "CIS605 Theme Park"
        'Using the theme park already created - disregard this testing entry !!
        'themePark = New ThemePark("CIS605 Theme Park")

        'Create feature, ID “F01”, Description “Park Pass”, Units “Day”, Adult Price $100, Child Price $80
        featureF01 = mThemePark.addFeature("F01", "Park Pass", "Day", 100, 80)

        'Create Feature, ID “F02”, Description “Early Entry Pass”, Units “Day”, Adult Price $10, Child Price $5
        featureF02 = mThemePark.addFeature("F02", "Early Entry Pass", "Day", 10, 5)

        'Create Feature, ID “F03”, Description “Meal Plan”, Units “Meal”, Adult Price $30, Child Price $20
        featureF03 = mThemePark.addFeature("F03", "Meal Plan", "Meal", 30, 20)

        'Create Customer, ID “C01”, Name “CName01”
        customerC01 = mThemePark.addCustomer("C01", "CName01")

        'Create Customer, ID “C02”, Name “CName02”
        customerC02 = mThemePark.addCustomer("C02", "CName02")

        'Create Customer, ID “C03”, Name “Customer Name 03”
        customerC03 = mThemePark.addCustomer("C03", "Customer Name 03")

        'Create Passbook, ID “PB01”, Customer “C01” reference, DatePurch 9/15/2015, Visitor Name “self”, Visitor BDay 1/1/1980
        passbookPB01 = mThemePark.addPassbook("PB01", customerC01, #9/15/2015#, "Self", #1/1/1980#)

        'Create Passbook, ID “PB02”, Customer “C02” reference, DatePurch 9/16/2015, Visitor Name “self”, Visitor BDay 6/1/1985
        passbookPB02 = mThemePark.addPassbook("PB02", customerC02, #9/16/2015#, "Self", #6/1/1985#)

        'Create Passbook, ID “PB03”, Customer “C02” reference, DatePurch 9/17/2015, Visitor Name “C02 Visitor”, Visitor BDay 12/1/2003
        passbookPB03 = mThemePark.addPassbook("PB03", customerC02, #9/17/2015#, "C02 Visitor", #12/1/2003#)

        'Create Passbook, ID “PB04”, Customer “C03” reference, DatePurch 8/15/2015, Visitor Name “self”, Visitor BDay 1/1/1975
        passbookPB04 = mThemePark.addPassbook("PB04", customerC03, #8/15/2015#, "Self", #1/1/1975#)

        'Create Passbook, ID “PB05”, Customer “C03” reference, DatePurch 9/15/2015, Visitor Name “C03 Visitor 1”, Visitor BDay 10/7/2002
        passbookPB05 = mThemePark.addPassbook("PB05", customerC03, #9/15/2015#, "Self", #10/7/2002#)

        'Create Passbook, ID “PB06”, Customer “C03” reference, DatePurch 10/15/2015, Visitor Name “C03 Visitor 2”, Visitor BDay 10/8/2002
        passbookPB06 = mThemePark.addPassbook("PB06", customerC03, #10/15/2015#, "Self", #10/8/2002#)

        'Purchase Feature (i.e. Create PassbookFeature), ID “PBF01”, Qty 1, Passbook “PB01” reference, feature “F01” reference
        pbf01 = mThemePark.addPassbookFeature("PBF01", 1, 85, passbookPB01, featureF01, 1)

        'Purchase Feature (i.e. Create PassbookFeature), ID “PBF02”, Qty 2, Passbook “PB02” reference, feature “F01” reference
        pbf02 = mThemePark.addPassbookFeature("PBF02", 2, 85, passbookPB02, featureF01, 1)

        'Purchase Feature (i.e. Create PassbookFeature), ID “PBF03”, Qty 3, Passbook “PB03” reference, feature “F01” reference
        pbf03 = mThemePark.addPassbookFeature("PBF03", 3, 85, passbookPB03, featureF01, 1)

        'Purchase Feature (i.e. Create PassbookFeature), ID “PBF04”, Qty 1, Passbook “PB04” reference, feature “F01” reference
        pbf04 = mThemePark.addPassbookFeature("PBF04", 1, 85, passbookPB04, featureF01, 1)

        'Purchase Feature (i.e. Create PassbookFeature), ID “PBF05”, Qty 1, Passbook “PB05” reference, feature “F01” reference
        pbf05 = mThemePark.addPassbookFeature("PBF05", 1, 85, passbookPB05, featureF01, 1)

        'Purchase Feature (i.e. Create PassbookFeature), ID “PBF06”, Qty 1, Passbook “PB06” reference, feature “F01” reference
        pbf06 = mThemePark.addPassbookFeature("PBF06", 1, 85, passbookPB06, featureF01, 1)

        'Purchase Feature (i.e. Create PassbookFeature), ID “PBF07”, Qty 3, Passbook “PB03” reference, feature “F02” reference
        pbf07 = mThemePark.addPassbookFeature("PBF07", 3, 85, passbookPB03, featureF02, 1)

        'Purchase Feature (i.e. Create PassbookFeature), ID “PBF08”, Qty 9, Passbook “PB03” reference, feature “F03” reference
        pbf08 = mThemePark.addPassbookFeature("PBF08", 9, 85, passbookPB03, featureF03, 1)

        'Purchase Feature (i.e. Create PassbookFeature), ID “PBF09”, Qty 1, Passbook “PB04” reference, feature “F01” reference
        pbf09 = mThemePark.addPassbookFeature("PBF09", 1, 85, passbookPB04, featureF01, 1)

        'Purchase Feature (i.e. Create PassbookFeature), ID “PBF10”, Qty 3, Passbook “PB04” reference, feature “F01” reference
        pbf10 = mThemePark.addPassbookFeature("PBF10", 3, 160, passbookPB04, featureF01, 1)

        'Use Feature (i.e. Create UsedFeature), ID “UF01”, PBFeature “PBF01” reference, DateUsed 10/20/2015, LocationUsed “Epcot Center”, QtyUsed 1
        uf01 = mThemePark.addUsedFeature("UF01", pbf01, #10/20/2015#, "Epcot Center", 1)

        'Use Feature(i.e.Create UsedFeature), ID “UF02”, PBFeature “PBF02” reference, DateUsed 20/20/2015, LocationUsed “West Parking”, QtyUsed 1
        'NOTE: I'm considering hte 20/20/2015 a typo as all the other are exact same date
        uf02 = mThemePark.addUsedFeature("UF02", pbf02, #10/20/2015#, "West Parking", 1)

        'Use Feature (i.e. Create UsedFeature), ID “UF03”, PBFeature “PBF03” reference, DateUsed 10/20/2015, LocationUsed “France”, QtyUsed 2
        uf01 = mThemePark.addUsedFeature("UF03", pbf03, #10/20/2015#, "France", 2)

        'Use Feature (i.e. Create UsedFeature), ID “UF04”, PBFeature “PBF03” reference, DateUsed 10/20/2015, LocationUsed “American Pavilion”, QtyUsed 1
        uf01 = mThemePark.addUsedFeature("UF04", pbf03, #10/20/2015#, "American Pavilion", 1)

        'Update Passbook Feature, PBFeatureID “PBF03”, DateUpdated 10/21/2015, QtyUpdated 1
        ' NOTE: Ideally this sends the feaure ID, but until we have storage of objects, sending reference
        mThemePark.updatePassbookFeature(pbf03, #10/21/2015#, 1)

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

    ''' <summary>
    ''' Automagically scrolls the transaction log when new entry hits it
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub _txtTransactionLogChanged(
            ByVal sender As Object,
            ByVal e As EventArgs
            ) _
        Handles txtTrxLogTabLog.TextChanged

        txtTrxLogTabLog.SelectionStart = txtTrxLogTabLog.TextLength
        txtTrxLogTabLog.ScrollToCaret()

    End Sub '_txtTrxLogTabChanged()


    '********** Business Logic Event Procedures
    '             - Initiated as a result of business logic
    '               method(s) running

    ''' <summary>
    ''' Handles processing when a ThemePark_CustomerAdded event is raised
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub _customerAdded(
            ByVal sender As System.Object,
            ByVal e As System.EventArgs) _
        Handles _
            mThemePark.ThemePark_CustomerAdded

        'Declare variables
        Dim theThemePark_EventArgs_CustomerAdded As _
            ThemePark_EventArgs_CustomerAdded
        Dim theCustomer As Customer

        'Get / Validate data
        theThemePark_EventArgs_CustomerAdded =
            CType(
                e,
                ThemePark_EventArgs_CustomerAdded
                )
        theCustomer = theThemePark_EventArgs_CustomerAdded.customer

        'Do processing
        'Add info to dashboard
        lstCustomersTabDashboard.Items.Add(theCustomer.custID)

        'Change the count on dashboard
        txtNumCustomersTabDashboard.Text = mThemePark.numCustomers.ToString

        'Add to combo box on buy passbook tab
        cboPassbookOwnerTabPurchasePassbook.Items.Add(theCustomer.custID)

        'Add to the log
        txtTrxLogTabLog.Text &= vbCrLf & CType(TimeValue(CType(Now, String)), String) _
            & " - New Customer Added: " _
            & theCustomer.ToString _
            & vbCrLf

    End Sub '_customerAdded()

    ''' <summary>
    ''' Handles processing when a ThemePark_FeatureAdded event is raised
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub _featureAdded(
            ByVal sender As System.Object,
            ByVal e As System.EventArgs) _
        Handles _
            mThemePark.ThemePark_FeatureAdded

        'Declare variables
        Dim theThemePark_EventArgs_FeatureAdded As _
            ThemePark_EventArgs_FeatureAdded
        Dim theFeature As Feature

        'Get / Validate data
        theThemePark_EventArgs_FeatureAdded =
            CType(
                e,
                ThemePark_EventArgs_FeatureAdded
                )
        theFeature = theThemePark_EventArgs_FeatureAdded.feature

        'Do processing
        'Add info to dashboard
        lstFeaturesTabDashboard.Items.Add(theFeature.featureID)

        'Update Feature Count with Newest Information
        txtNumFeaturesTabDashboard.Text = mThemePark.numFeatures.ToString

        'Add to combo box on buy feature tab
        cboFeatureSelectTabBuyFeature.Items.Add(theFeature.featureID)

        'Add to combo box on post used feature tab
        cboFeatureSelectorTabPostUsedFeature.Items.Add(theFeature.featureID)

        'Add to the transaction log
        txtTrxLogTabLog.Text &= vbCrLf & CType(TimeValue(CType(Now, String)), String) _
            & " - New Feature Added: " _
            & theFeature.ToString _
            & vbCrLf

    End Sub '_featureAdded()

    ''' <summary>
    ''' Handles processing when a ThemePark_UsedFeatureAdded event is raised
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub _usedFeatureAdded(
            ByVal sender As System.Object,
            ByVal e As System.EventArgs) _
        Handles _
            mThemePark.ThemePark_UsedFeatureAdded

        'Declare variables
        Dim theThemePark_EventArgs_UsedFeatureAdded As _
            ThemePark_EventArgs_UsedFeatureAdded
        Dim theUsedFeature As UsedFeature

        'Get / Validate data
        theThemePark_EventArgs_UsedFeatureAdded =
            CType(
                e,
                ThemePark_EventArgs_UsedFeatureAdded
                )
        theUsedFeature = theThemePark_EventArgs_UsedFeatureAdded.usedFeature

        'Do processing
        'Update Dashboard & Feature Count with Newest Information
        lstUsedFeatureTabDashboard.Items.Add(theUsedFeature.usedFeatureID.ToString)
        txtNumFeaturesTabDashboard.Text = mThemePark.numUsedFeatures.ToString

        'Add to the transaction log
        txtTrxLogTabLog.Text &= vbCrLf & CType(TimeValue(CType(Now, String)), String) _
            & " - Used Feature Added: " _
            & theUsedFeature.ToString _
            & vbCrLf

    End Sub '_usedFeatureAdded()

    ''' <summary>
    ''' Event Handler for ThemePark_PassbookAdded Event
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub _passbookAdded(
            ByVal sender As System.Object,
            ByVal e As System.EventArgs) _
        Handles _
            mThemePark.ThemePark_PassbookAdded

        'Declare variables
        Dim theThemePark_EventArgs_Passbook_Added As _
            ThemePark_EventArgs_PassbookAdded
        Dim thePassbook As Passbook

        'Get / Validate data
        theThemePark_EventArgs_Passbook_Added =
            CType(
                e,
                ThemePark_EventArgs_PassbookAdded
                )
        thePassbook = theThemePark_EventArgs_Passbook_Added.passbook

        'Do Processing
        'Add info to dashboard
        lstPassbooksTabDashboard.Items.Add(thePassbook.passbookID)

        'Change the count on the dashboard
        txtNumPassbooksTabDashboard.Text = mThemePark.numPassbooks.ToString

        'Add To Combo Box on add feature to passbook tab and use feature tab
        cboPassbookIDTabPostUsedFeature.Items.Add(thePassbook.passbookID)
        cboPassbookTabUpdatePassbook.Items.Add(thePassbook.passbookID)
        cboPickPassbookIDTabBuyFeature.Items.Add(thePassbook.passbookID)

        'Add to the log
        txtTrxLogTabLog.Text &= vbCrLf & CType(TimeValue(CType(Now, String)), String) _
            & " - New Passbook Added: " _
            & thePassbook.ToString _
            & vbCrLf

    End Sub '_passbookAdded()

    ''' <summary>
    ''' Event handler for ThemePark_PassbookFeatureAdded events
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub _passbookFeatureAdded(
            ByVal sender As System.Object,
            ByVal e As System.EventArgs) _
        Handles _
            mThemePark.ThemePark_PassbookFeatureAdded

        'Declare the variables
        Dim theThemePark_EventArgs_PassbookFeatureAdded As _
            ThemePark_EventArgs_PassbookFeatureAdded
        Dim thePassbookFeature As PassbookFeature

        'Get / Validate the data
        theThemePark_EventArgs_PassbookFeatureAdded =
            CType(
                e,
                ThemePark_EventArgs_PassbookFeatureAdded
                )
        thePassbookFeature = theThemePark_EventArgs_PassbookFeatureAdded.passbookFeature

        'Do processing
        'Add info to Dashboard
        lstPassbookFeatureTabDashboard.Items.Add(thePassbookFeature.passbookFeatureID)

        'change the count on the dashboard
        txtNumPassbookFeatureTabDashboard.Text = mThemePark.numPassbookFeatures.ToString

        'Add to the combo boxes
        cboFeatureSelectorTabPostUsedFeature.Items.Add(thePassbookFeature.passbookFeatureID)

        'Add to the log
        txtTrxLogTabLog.Text &= vbCrLf & CType(TimeValue(CType(Now, String)), String) _
            & " - New Passbook Feature Added: " _
            & thePassbookFeature.ToString _
            & vbCrLf

    End Sub '_passbookFeatureAdded()



#End Region 'Event Procedures

#Region "Events"
    '******************************************************************
    'Events
    '******************************************************************

    'No Events are currently defined.
    'These are all public.

#End Region 'Events 

End Class 'FrmMain
