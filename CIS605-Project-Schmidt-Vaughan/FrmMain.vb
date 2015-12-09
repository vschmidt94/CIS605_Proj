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
Imports System.IO
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

    ''' <summary>
    ''' Helper function that will clean the dashboard
    ''' </summary>
    Sub _cleanDashboard(ByVal pList As ListBox)

        ' reset the selected features across all lists
        If pList IsNot lstCustomersTabDashboard Then
            lstCustomersTabDashboard.SelectedIndex = -1
        End If

        If pList IsNot lstFeaturesTabDashboard Then
            lstFeaturesTabDashboard.SelectedIndex = -1
        End If

        If pList IsNot lstPassbookFeatureTabDashboard Then
            lstPassbookFeatureTabDashboard.SelectedIndex = -1
        End If

        If pList IsNot lstPassbooksTabDashboard Then
            lstPassbooksTabDashboard.SelectedIndex = -1
        End If

        If pList IsNot lstUsedFeatureTabDashboard Then
            lstUsedFeatureTabDashboard.SelectedIndex = -1
        End If

    End Sub 'cleanDashboard()

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

        'Look to see if customer ID already exists. Duplicate customers not allowed
        If mThemePark.findCustomer(_custID) >= 0 Then
            MessageBox.Show("Customer ID already exists!")
            txtAddCustIDTabNewCustomer.SelectAll()
            txtAddCustIDTabNewCustomer.Focus()
            Exit Sub
        End If

        'Create a new Customer
        mThemePark.addCustomer(_custID, _custName)

        'Confirm with message box and clear the form for more
        MessageBox.Show("New Customer Added")
        txtAddCustIDTabNewCustomer.ResetText()
        txtAddCustNameTabNewCustomer.ResetText()

    End Sub 'btnAddCustomer_Click()

    ''' <summary>
    ''' Clears the text fields on button click
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnClearTabNewCustomer_Click(
            sender As Object,
            e As EventArgs) _
        Handles btnClearTabNewCustomer.Click

        'Clear the fields
        txtAddCustIDTabNewCustomer.ResetText()
        txtAddCustNameTabNewCustomer.ResetText()

    End Sub 'btnClearTabNewCustomer_Click

    Private Sub btnAddFeature_Click(
            sender As Object,
            e As EventArgs
            ) _
        Handles btnAddFeatureTabDefineFeature.Click

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


        'Look to see if Feature ID already exists. Duplicate features not allowed
        If mThemePark.findFeature(_featureID) >= 0 Then
            MessageBox.Show("Feature ID already exists!")
            txtNewFeatureIDTabDefineFeature.SelectAll()
            txtNewFeatureIDTabDefineFeature.Focus()
            Exit Sub
        End If


        'Create New Feature
        mThemePark.addFeature(_featureID,
                              _featureName,
                              _featureUOM,
                              _featureAdultPrice,
                              _featureChildPrice)

        'Confirm with message box and clear the form for more
        MessageBox.Show("New Feature Added")
        txtNewFeatureNameTabDefineFeature.ResetText()
        txtNewFeatureIDTabDefineFeature.ResetText()
        txtFeatureUOMTabDefineFeature.ResetText()
        txtFeatureAdultPriceTabDefineFeature.ResetText()
        txtFeatureChildPriceTabDefineFeature.ResetText()

    End Sub 'btnAddFeature_Click() 

    ''' <summary>
    ''' Purchases a new passbook
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnBuyPassbookTabPurhchasePassbook_Click(
            sender As Object,
            e As EventArgs) _
        Handles btnBuyPassbookTabPurhchasePassbook.Click

        'Declare Variables
        Dim _passbookID As String
        Dim _passbookOwner As Customer
        Dim _passbookUser As String
        Dim _passbookPurchDate As Date
        Dim _passbookUserBirthDate As Date

        'Validate information on the form
        If txtPassbookIDTabPurchasePassbook.Text.Length = 0 Then
            MessageBox.Show("Passbook ID must be entered!")
            txtPassbookIDTabPurchasePassbook.Focus()
            Exit Sub
        End If

        If txtPassbookUserTabPurchasePassbook.Text.Length = 0 Then
            MessageBox.Show("Passbook user must be entered!")
            txtPassbookUserTabPurchasePassbook.Focus()
            Exit Sub
        End If

        If cboPassbookOwnerTabPurchasePassbook.SelectedIndex < 0 Then
            MessageBox.Show("Passbook owner must be selected!")
            cboPassbookOwnerTabPurchasePassbook.Focus()
            Exit Sub
        End If

        'Check if Passbook already exists
        _passbookID = txtPassbookIDTabPurchasePassbook.Text
        If mThemePark.findPassbook(_passbookID) >= 0 Then
            MessageBox.Show("Passbook ID Already Exists!")
            txtPassbookIDTabPurchasePassbook.SelectAll()
            txtPassbookIDTabPurchasePassbook.Focus()
            Exit Sub
        End If

        'If here, we should have at least nominally OK data on 
        'form. Load the form data into local variables
        _passbookOwner = mThemePark.ithCustomer(cboPassbookOwnerTabPurchasePassbook.SelectedIndex)
        _passbookUser = txtPassbookUserTabPurchasePassbook.Text
        _passbookPurchDate = dtpPassbookPurchDateTabPurchasePassbook.Value
        _passbookUserBirthDate = dtpPassbookUserBirthdateTabPurchasePassbook.Value

        'Add new passbook to themePark
        mThemePark.addPassbook(_passbookID,
                                _passbookOwner,
                                _passbookPurchDate,
                                _passbookUser,
                                _passbookUserBirthDate)

        'Confirmation message
        MessageBox.Show("Passbook Added!")

        'Clean up
        txtPassbookIDTabPurchasePassbook.ResetText()
        txtPassbookUserTabPurchasePassbook.ResetText()
        cboPassbookOwnerTabPurchasePassbook.SelectedIndex = -1
        dtpPassbookPurchDateTabPurchasePassbook.ResetText()
        dtpPassbookUserBirthdateTabPurchasePassbook.ResetText()

    End Sub 'btnBuyPassbookTabPurchasePassbook_Click(sender As Object, e As EventArgs) Handles btnBuyPassbookTabPurhchasePassbook.Click

    ''' <summary>
    ''' Updates the selected customer text box if needed
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub lstCustomersTabDashboard_SelectedIdxChanged(
                sender As Object,
                e As EventArgs) _
        Handles lstCustomersTabDashboard.SelectedIndexChanged

        'Need to bail if the new selected index is -1
        If lstCustomersTabDashboard.SelectedIndex = -1 Then
            Exit Sub
        End If

        'call a helper function that cleans up the dashboard
        _cleanDashboard(lstCustomersTabDashboard)

        'display details about particular customer
        txtDetailsTabDashboard.Text =
            mThemePark.ithCustomer(lstCustomersTabDashboard.SelectedIndex).ToString

    End Sub 'lstCustomersTabDashboard_SelectedIdxChanged()

    Private Sub lstFeaturesTabDashboard_SelectedIdxChanged(sender As Object, e As EventArgs) _
        Handles lstFeaturesTabDashboard.SelectedIndexChanged

        'Need to bail if the new selected index is -1
        If lstFeaturesTabDashboard.SelectedIndex = -1 Then
            Exit Sub
        End If

        'call a helper function that cleans up the dashboard
        _cleanDashboard(lstFeaturesTabDashboard)

        'display details about particular customer
        txtDetailsTabDashboard.Text =
            mThemePark.ithFeature(lstFeaturesTabDashboard.SelectedIndex).ToString

    End Sub 'lstFeaturesTabDashboard_SelectedIdxChanged()

    Private Sub lstPassbooksTabDashboard_SelectedIdxChanged(sender As Object, e As EventArgs) _
        Handles lstPassbooksTabDashboard.SelectedIndexChanged

        'Need to bail if the new selected index is -1
        If lstPassbooksTabDashboard.SelectedIndex = -1 Then
            Exit Sub
        End If

        'call a helper function that cleans up the dashboard
        _cleanDashboard(lstPassbooksTabDashboard)

        'display details about particular customer
        txtDetailsTabDashboard.Text =
            mThemePark.ithPassbook(lstPassbooksTabDashboard.SelectedIndex).ToString

    End Sub 'lstPassbooksTabDashboard_SelectedIdxChanged()

    Private Sub lstPassbookFeatureTabDashboard_SelectedIdxChanged(sender As Object, e As EventArgs) _
        Handles lstPassbookFeatureTabDashboard.SelectedIndexChanged

        'Need to bail if the new selected index is -1
        If lstPassbookFeatureTabDashboard.SelectedIndex = -1 Then
            Exit Sub
        End If

        'call a helper function that cleans up the dashboard
        _cleanDashboard(lstPassbookFeatureTabDashboard)

        'display details about particular customer
        txtDetailsTabDashboard.Text =
            mThemePark.ithPassbookFeature(lstPassbookFeatureTabDashboard.SelectedIndex).ToString

    End Sub 'lstPassbookFeatureTabDashboard_SelectedIdxChanged()

    Private Sub lstUsedFeatureTabDashboard_SelectedIdxChanged(sender As Object, e As EventArgs) _
        Handles lstUsedFeatureTabDashboard.SelectedIndexChanged

        'Need to bail if the new selected index is -1
        If lstUsedFeatureTabDashboard.SelectedIndex = -1 Then
            Exit Sub
        End If

        'call a helper function that cleans up the dashboard
        _cleanDashboard(lstUsedFeatureTabDashboard)

        'display details about particular customer
        txtDetailsTabDashboard.Text =
            mThemePark.ithUsedFeature(lstUsedFeatureTabDashboard.SelectedIndex).ToString

    End Sub 'lstPassbookFeatureTabDashboard_SelectedIdxChanged()

    ''' <summary>
    ''' Displays the selected customer in the text box on purchase passbook tab
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub cboPassbookOwnerTabPurchasePassbook_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboPassbookOwnerTabPurchasePassbook.SelectedIndexChanged

        If cboPassbookOwnerTabPurchasePassbook.SelectedIndex < 0 Then
            'If the selected index is -1, then just clear the form
            txtCustDetailsTabBuyPassbook.ResetText()
        Else
            'Update the text in the info text box
            txtCustDetailsTabBuyPassbook.Text =
            mThemePark.ithCustomer(cboPassbookOwnerTabPurchasePassbook.SelectedIndex).ToString
        End If

    End Sub 'cboPassbookOwnerTabPurchasePassbook_SelectedIndexChanged()

    ''' <summary>
    ''' Populate form based on selected feature in combobox / buy feature tab
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub cboPickPassbookIDTabBuyFeature_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboPickPassbookIDTabBuyFeature.SelectedIndexChanged

        'Declare variables
        Dim selectedIdx As Integer = cboPickPassbookIDTabBuyFeature.SelectedIndex
        Dim cost As Decimal

        'Bail if not a valid index
        If selectedIdx < 0 Then
            Exit Sub
        End If

        'Put Passbook string info in text box
        txtPassbookStringTabBuyFeature.Text =
            mThemePark.ithPassbook(selectedIdx).ToString

        'Populate Registered Owner and User, Age etc.
        txtRegisteredOwnerTabBuyFeature.Text =
            mThemePark.ithPassbook(selectedIdx).passbookOwner.custName
        txtRegisteredUserTabBuyFeature.Text =
            mThemePark.ithPassbook(selectedIdx).passbookVisitorName
        txtUserAgeTabBuyFeature.Text =
            mThemePark.ithPassbook(selectedIdx).age.ToString
        chkIsChildTabBuyFeature.Checked =
            mThemePark.ithPassbook(selectedIdx).isChild

        'if we have already selected a feature, need to update the price
        If cboFeatureSelectTabBuyFeature.SelectedIndex >= 0 Then
            If chkIsChildTabBuyFeature.Checked Then
                cost =
                    mThemePark.ithFeature(cboFeatureSelectTabBuyFeature.SelectedIndex).featureChildPrice
            Else
                cost =
                    mThemePark.ithFeature(cboFeatureSelectTabBuyFeature.SelectedIndex).featureAdultPrice
            End If

            cost = cost * numQtyTabBuyFeature.Value

            txtTotalCostTabBuyFeature.Text =
                cost.ToString("C0")
        End If


    End Sub 'cboPickPassbookIDTabBuyFeature_SelectedIndexChanged()

    ''' <summary>
    ''' Populate the right fields when a feature is selected from combo box
    ''' on Buy Feature tab
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub cboFeatureSelectTabBuyFeature_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboFeatureSelectTabBuyFeature.SelectedIndexChanged

        'Declare variables
        Dim selectedIdx As Integer = cboFeatureSelectTabBuyFeature.SelectedIndex
        Dim cost As Decimal

        'If the feature index is not valid, bail out now
        If selectedIdx < 0 Then
            Exit Sub
        End If

        'populate fields with the information
        If chkIsChildTabBuyFeature.Checked Then
            txtFeaturePricePerUnitTabBuyFeature.Text =
                mThemePark.ithFeature(selectedIdx).featureChildPrice.ToString("C2")
            cost = mThemePark.ithFeature(selectedIdx).featureChildPrice
        Else
            txtFeaturePricePerUnitTabBuyFeature.Text =
                mThemePark.ithFeature(selectedIdx).featureAdultPrice.ToString("C2")
            cost = mThemePark.ithFeature(selectedIdx).featureAdultPrice
        End If

        txtFeatureUOMTabBuyFeature.Text =
            mThemePark.ithFeature(selectedIdx).featureUOM

        txtFeatureStringTabBuyFeature.Text =
            mThemePark.ithFeature(selectedIdx).ToString

        cost = cost * numQtyTabBuyFeature.Value

        txtTotalCostTabBuyFeature.Text = cost.ToString("C0")

    End Sub 'cboFeatureSelectTabBuyFeature_SelectedIndexChanged()

    ''' <summary>
    ''' Resets the buy feature tab when cancel button is pressed.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnCancelTabBuyFeature_Click(sender As Object, e As EventArgs) Handles btnCancelTabBuyFeature.Click

        'Clear text fields
        txtPassbookStringTabBuyFeature.ResetText()
        txtRegisteredOwnerTabBuyFeature.ResetText()
        txtRegisteredUserTabBuyFeature.ResetText()
        txtPassbookUserTabPurchasePassbook.ResetText()
        txtFeatureStringTabBuyFeature.ResetText()
        txtUserAgeTabBuyFeature.ResetText()
        txtFeatureUOMTabBuyFeature.ResetText()
        txtFeaturePricePerUnitTabBuyFeature.ResetText()
        txtTotalCostTabBuyFeature.ResetText()


        'Reset combo boxes, etc.
        numQtyTabBuyFeature.Value = 1
        cboFeatureSelectTabBuyFeature.SelectedIndex = -1
        cboFeatureSelectTabBuyFeature.Text = "Select Feature"
        cboPickPassbookIDTabBuyFeature.SelectedIndex = -1
        cboPickPassbookIDTabBuyFeature.Text = "Select Passbook"

    End Sub 'btnCancelTabBuyFeature_Click()

    ''' <summary>
    ''' Updates the total cost if needed on changed qty
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub numQtyTabBuyFeature_ValueChanged(
            sender As Object,
            e As EventArgs) _
        Handles numQtyTabBuyFeature.ValueChanged

        'Declare variables
        Dim cost As Decimal

        'Only need to update the cost if a passbook and feature are selected
        'So, if these don't have selected indexes >= 0 just bail out,
        'nothing to do
        If cboPickPassbookIDTabBuyFeature.SelectedIndex < 0 Then
            Exit Sub
        End If

        If cboFeatureSelectTabBuyFeature.SelectedIndex < 0 Then
            Exit Sub
        End If

        'Get the right unit cost based on age
        If chkIsChildTabBuyFeature.Checked Then
            cost = mThemePark.ithFeature(cboFeatureSelectTabBuyFeature.SelectedIndex).featureChildPrice
        Else
            cost = mThemePark.ithFeature(cboFeatureSelectTabBuyFeature.SelectedIndex).featureAdultPrice
        End If

        'extend cost based on qty
        cost = cost * numQtyTabBuyFeature.Value

        'Update text
        txtTotalCostTabBuyFeature.Text = cost.ToString("C2")

    End Sub 'numQtyTabBuyFeature_ValueChanged()

    ''' <summary>
    ''' Purchases a feature for a passbook
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnBuyTabBuyFeature_Click(
            sender As Object,
            e As EventArgs) _
        Handles btnBuyTabBuyFeature.Click

        'Declare variables
        Dim thePassbook As Passbook
        Dim theFeature As Feature
        Dim theQty As Decimal
        Dim thePurchasePrice As Decimal
        Dim theNewID As String

        'Parse the form, validate we have enough stuff
        'to properly form a new object

        If cboPickPassbookIDTabBuyFeature.SelectedIndex < 0 Then
            MessageBox.Show("Must select a passbook!")
            cboPickPassbookIDTabBuyFeature.Focus()
            Exit Sub
        End If

        thePassbook = mThemePark.ithPassbook(cboPickPassbookIDTabBuyFeature.SelectedIndex)

        If cboFeatureSelectTabBuyFeature.SelectedIndex < 0 Then
            MessageBox.Show("Must select a feature!")
            cboFeatureSelectTabBuyFeature.Focus()
            Exit Sub
        End If

        theFeature = mThemePark.ithFeature(cboFeatureSelectTabBuyFeature.SelectedIndex)

        theQty = numQtyTabBuyFeature.Value

        ' This is only checking for *something* in the new id field.
        ' TODO: improve robustness
        If txtNewPassbookFeatureIdTabBuyFeature.Text.Length < 1 Then
            MessageBox.Show("Must have new Feature ID")
            txtNewPassbookFeatureIdTabBuyFeature.Focus()
            Exit Sub
        End If

        'Calculate the correct purchase price
        If chkIsChildTabBuyFeature.Checked Then
            thePurchasePrice = mThemePark.ithFeature(cboFeatureSelectTabBuyFeature.SelectedIndex).featureChildPrice
        Else
            thePurchasePrice = mThemePark.ithFeature(cboFeatureSelectTabBuyFeature.SelectedIndex).featureAdultPrice
        End If
        thePurchasePrice = thePurchasePrice * theQty

        theNewID = txtNewPassbookFeatureIdTabBuyFeature.Text

        'Ensure we don't have a duplicate ID
        If mThemePark.findPassbookFeature(theNewID) >= 0 Then
            MessageBox.Show("Passbook Feature ID already exists!")
            txtNewPassbookFeatureIdTabBuyFeature.SelectAll()
            txtNewPassbookFeatureIdTabBuyFeature.Focus()
            Exit Sub
        End If

        'Add new passbook feature
        mThemePark.addPassbookFeature(theNewID, theQty, thePurchasePrice, thePassbook, theFeature, theQty)

        'Display message on success
        MessageBox.Show("Feature Added to Passbook")

    End Sub 'btnBuyTabBuyFeature_Clicked()

    ''' <summary>
    ''' Updates the form on change of selected passbook id
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub cboPassbookIDTabPostUsedFeature_SelectedIndexChanged(
            sender As Object,
            e As EventArgs) _
        Handles cboPassbookIDTabPostUsedFeature.SelectedIndexChanged,
                cboFeatureSelectorTabPostUsedFeature.SelectedIndexChanged

        'Declare variables
        Dim thePassbookFeature As PassbookFeature
        Dim thePBFeatureID As String
        Dim thePassbookID As String
        Dim theFeatureID As String
        Dim i As Integer

        'If the selected passbook index is invalid, bail out now
        If cboPassbookIDTabPostUsedFeature.SelectedIndex < 0 Then
            Exit Sub
        End If

        'If there are no passbook features, bail out now.
        If mThemePark.numPassbookFeatures < 1 Then
            Exit Sub
        End If

        'Update text fields
        txtPassbookOwnerTabPostUsedFeature.Text =
            mThemePark.ithPassbook(cboPassbookIDTabPostUsedFeature.SelectedIndex).passbookOwner.custName
        txtPassbookUserTabPostUsedFeature.Text =
            mThemePark.ithPassbook(cboPassbookIDTabPostUsedFeature.SelectedIndex).passbookVisitorName

        'See if there is a feature selected, if not we've updated
        'everything we can for now so bail out.
        If cboFeatureSelectorTabPostUsedFeature.SelectedIndex < 0 Then
            Exit Sub
        End If

        'OK, we have enough stuff to look and see if there is
        'a matching passbook feature.
        'First, get the stuff we want to search for in local 
        'variables for convenience
        thePassbookID =
            mThemePark.ithPassbook(cboPassbookIDTabPostUsedFeature.SelectedIndex).passbookID
        theFeatureID =
            mThemePark.ithFeature(cboFeatureSelectorTabPostUsedFeature.SelectedIndex).featureID

        'Loop through the Passbook features to see if we have one that matches
        For i = 0 To mThemePark.numPassbookFeatures - 1
            If mThemePark.ithPassbookFeature(i).passbook.passbookID = thePassbookID And
                    mThemePark.ithPassbookFeature(i).feature.featureID = theFeatureID Then
                'We found a match
                thePassbookFeature = mThemePark.ithPassbookFeature(i)
                Exit For
            End If
        Next

        'if we didn't find a match, nothing more can be done
        If thePassbookFeature Is Nothing Then
            MessageBox.Show("That feature does not exist in the choosen passbook")

            'Clear out any text values that might be remaining
            txtRemainingQtyTabPostUsedFeature.Text = "N/A"

            'Disable the update button for now
            btnUpdateTabPostUsedFeature.Enabled = False
            Exit Sub
        End If

        'if here we must have something
        'update the text fields
        txtRemainingQtyTabPostUsedFeature.Text = thePassbookFeature.passbookFeatureQtyRemaining.ToString

        'make sure the update button is enabled
        btnUpdateTabPostUsedFeature.Enabled = True

    End Sub 'cboPassbookIDTabPostUsedFeature_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboPassbookIDTabPostUsedFeature.SelectedIndexChanged)

    ''' <summary>
    ''' Adds a used feature on button click on Post Used Feature tab
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnUpdateTabPostUsedFeature_Click(
            sender As Object,
            e As EventArgs) _
        Handles btnUpdateTabPostUsedFeature.Click

        'Declare variables
        Dim thePassbookID As String
        Dim theFeatureID As String
        Dim i As Integer
        Dim thePassbookFeature As PassbookFeature

        Dim theQtyRemaining As Decimal
        Dim theQtyUsed As Decimal

        'Our update passbook button should only enable for good combinations
        'of Passbooks and Passbook Features
        'So, if we're here we know we should have a good combination

        'Make sure the feature exists in the passbook and has qty remaining
        'First, get the stuff we want to search for in local 
        'variables for convenience
        thePassbookID =
            mThemePark.ithPassbook(cboPassbookIDTabPostUsedFeature.SelectedIndex).passbookID
        theFeatureID =
            mThemePark.ithFeature(cboFeatureSelectorTabPostUsedFeature.SelectedIndex).featureID

        'Loop through the Passbook features to find the match if we have one that matches
        For i = 0 To mThemePark.numPassbookFeatures - 1
            If mThemePark.ithPassbookFeature(i).passbook.passbookID = thePassbookID And
                    mThemePark.ithPassbookFeature(i).feature.featureID = theFeatureID Then
                'We found a match
                thePassbookFeature = mThemePark.ithPassbookFeature(i)
                Exit For
            End If
        Next

        'Make sure we have a valid qty used
        Try
            theQtyUsed = Decimal.Parse(txtUsedQtyTabPostUsedFeature.Text)
        Catch ex As Exception
            MessageBox.Show("Please enter a valid decimal value (example: 1.23)")
            txtUsedQtyTabPostUsedFeature.SelectAll()
            txtUsedQtyTabPostUsedFeature.Focus()
            txtTrxLogTabLog.Text &= vbCrLf & "Invalid Qty Used Entered"
            Exit Sub
        End Try

        theQtyRemaining = thePassbookFeature.passbookFeatureQtyRemaining()
        If theQtyRemaining <= 0 Then
            MessageBox.Show("There is no qty remaining for this passbook feature! Nothing can be updated!")
            Exit Sub
        End If

        'Make sure we aren't using more than we have available
        If theQtyUsed > theQtyRemaining Then
            MessageBox.Show("Qty Used exceeds Qty Available!")
            txtUsedQtyTabPostUsedFeature.SelectAll()
            txtUsedQtyTabPostUsedFeature.Focus()
            Exit Sub
        End If

        'Make sure we have a new feature ID
        If txtNewUsedFeatureIDTabPostUsedFeature.Text.Length < 1 Then
            MessageBox.Show("Must enter a new used feature ID")
            txtNewUsedFeatureIDTabPostUsedFeature.Focus()
            Exit Sub
        End If

        'Make sure we have a location used
        If txtLocationUsedTabPostUsedFeature.Text.Length < 3 Then
            MessageBox.Show("Must have a decent location used string (at least 3 chars).")
            txtLocationUsedTabPostUsedFeature.SelectAll()
            txtLocationUsedTabPostUsedFeature.Focus()
            Exit Sub
        End If

        'If here, good to proceed with updating records.
        mThemePark.addUsedFeature(
            txtNewUsedFeatureIDTabPostUsedFeature.Text,
            thePassbookFeature,
            dtpUsedDateTabPostUsedFeature.Value,
            txtLocationUsedTabPostUsedFeature.Text,
            theQtyUsed)

        'Display success message
        MessageBox.Show("Successfully added used feature!")

        'Update the form to show new qty remaining
        txtRemainingQtyTabPostUsedFeature.Text = thePassbookFeature.passbookFeatureQtyRemaining.ToString

        'Clear out the old entry
        txtUsedQtyTabPostUsedFeature.ResetText()

    End Sub 'btnUpdateTabPostUsedFeature_Click(sender As Object, e As EventArgs) Handles btnUpdateTabPostUsedFeature.Click

    Private Sub cboPassbookTabUpdatePassbook_SelectedIndexChanged(
            sender As Object,
            e As EventArgs) _
        Handles cboPassbookTabUpdatePassbook.SelectedIndexChanged

        'Declare variables
        Dim thePassbook As Passbook
        Dim thePassbookFeature As PassbookFeature
        Dim theFeature As Feature
        Dim theUsedFeature As UsedFeature
        Dim i As Integer
        Dim j As Integer

        txtNewQtyRemainingTabUpdatePassbook.ResetText()

        'Need to bail if the new selected index is -1
        If cboPassbookTabUpdatePassbook.SelectedIndex = -1 Then
            Exit Sub
        End If

        'Update the text boxes
        thePassbook = mThemePark.ithPassbook(cboPassbookTabUpdatePassbook.SelectedIndex)
        txtRegisteredOwnerTabUpdatePassbook.Text = thePassbook.passbookOwner.custName
        txtRegisteredUserTabUpdatePassbook.Text = thePassbook.passbookVisitorName

        txtAgeBoxTabUpdatePassbook.Text = thePassbook.age.ToString
        chkUserIsChildTabUpdatePassbook.Checked = thePassbook.isChild

        'Loop Through the passbook features and populate the list if it was in this passbook
        lstRemainFeatNameTabUpdatePassbook.Items.Clear()
        lstQtyRemainingTabUpdatePassbook.Items.Clear()
        lstFeatureUpdateTabUpdatePassbook.Items.Clear()
        For i = 0 To mThemePark.numPassbookFeatures - 1
            thePassbookFeature = mThemePark.ithPassbookFeature(i)
            If thePassbookFeature.passbook.passbookID = thePassbook.passbookID Then
                'we have a match
                lstFeatureUpdateTabUpdatePassbook.Items.Add(thePassbookFeature.passbookFeatureID)
                lstQtyRemainingTabUpdatePassbook.Items.Add(thePassbookFeature.passbookFeatureQtyRemaining.ToString("N2"))
                'Loop throught the features to get the name
                For j = 0 To mThemePark.numFeatures - 1
                    theFeature = mThemePark.ithFeature(j)
                    If thePassbookFeature.feature.featureID = theFeature.featureID Then
                        lstRemainFeatNameTabUpdatePassbook.Items.Add(theFeature.featureName)
                    End If
                Next
            End If
        Next

        'Loop through the used feature and populate that list if it was from the passbook
        lstUsedFeatNameTabUpdatePassbook.Items.Clear()
        lstUsedFeatTabUpdatePassbook.Items.Clear()
        lstQtyUsedTabUpdatePassbook.Items.Clear()
        lstLocUsedTabUpdatePassbook.Items.Clear()
        For i = 0 To mThemePark.numUsedFeatures - 1
            theUsedFeature = mThemePark.ithUsedFeature(i)
            If theUsedFeature.usedPassbookFeature.passbook.passbookID = thePassbook.passbookID Then
                'we have a match
                lstUsedFeatTabUpdatePassbook.Items.Add(theUsedFeature.usedFeatureID)
                lstQtyUsedTabUpdatePassbook.Items.Add(theUsedFeature.usedQty.ToString("N2"))
                lstLocUsedTabUpdatePassbook.Items.Add(theUsedFeature.usedLocation)
                'Loop through the features to get the name
                For j = 0 To mThemePark.numFeatures - 1
                    theFeature = mThemePark.ithFeature(j)
                    If theUsedFeature.usedPassbookFeature.feature.featureID = theFeature.featureID Then
                        lstUsedFeatNameTabUpdatePassbook.Items.Add(theFeature.featureName)
                    End If
                Next
            End If
        Next

    End Sub 'cboPassbookTabUpdatePassbook_SelectedIndexChanged(sender As Object, e As EventArgs)

    ''' <summary>
    ''' Keeps the items in sync on list list boxes on update passbook tab
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub lstFeatureUpdateTabUpdatePassbook_SelectedIndexChanged(
            sender As Object,
            e As EventArgs) _
        Handles lstFeatureUpdateTabUpdatePassbook.SelectedIndexChanged

        'Declare Variables
        Dim i As Integer

        'Bail if index was changed to -1
        If lstFeatureUpdateTabUpdatePassbook.SelectedIndex = -1 Then
            txtNewQtyRemainingTabUpdatePassbook.ResetText()
            Exit Sub
        End If

        'Sync the other list boxes
        lstRemainFeatNameTabUpdatePassbook.SelectedIndex = lstFeatureUpdateTabUpdatePassbook.SelectedIndex
        lstQtyRemainingTabUpdatePassbook.SelectedIndex = lstFeatureUpdateTabUpdatePassbook.SelectedIndex

        'Update the revised qty txt box with current qty
        txtNewQtyRemainingTabUpdatePassbook.Text = lstQtyRemainingTabUpdatePassbook.SelectedItem.ToString

    End Sub 'lstFeatureUpdateTabUpdatePassbook_SelectedIndexChanged()

    ''' <summary>
    ''' Updates the anticipated cost of the change to passbook on update passbook tab
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub txtNewQtyRemainingTabUpdatePassbook_TextChanged(
            sender As Object,
            e As EventArgs) _
        Handles txtNewQtyRemainingTabUpdatePassbook.TextChanged

        Dim theChangeCost As Decimal
        Dim theNewQtyRemaining As Decimal
        Dim theFeatureID As String


        Try
            theNewQtyRemaining = Decimal.Parse(txtNewQtyRemainingTabUpdatePassbook.Text)
        Catch ex As Exception
            MessageBox.Show("Please enter a valid decimal quantity (example: 1.23)")
            txtNewQtyRemainingTabUpdatePassbook.SelectAll()
            txtNewQtyRemainingTabUpdatePassbook.Focus()
            txtTrxLogTabLog.Text &= vbCrLf & "Invalid New Qty Remaining Value Entered"
            Exit Sub
        End Try

        If lstFeatureUpdateTabUpdatePassbook.SelectedIndex = -1 Then
            MessageBox.Show("Must select a feature to be changed!")
            Exit Sub
        End If

        theFeatureID = lstFeatureUpdateTabUpdatePassbook.SelectedItem.ToString



    End Sub 'txtNewQtyRemainingTabUpdatePassbook_TextChanged()

    Private Sub btnReadFile_Click(sender As Object, e As EventArgs) _
        Handles btnReadFile.Click

        Dim inputFile = New StreamReader("Data-in-single.txt")
        Dim outputFile = New StreamWriter("Data-out-single.txt")
        Dim idx As Integer
        Dim line As String
        Dim field() As String
        Dim fieldIdx As Integer
        Dim objectName As String
        Dim objectAction As String
        Dim objectTrxDate As String
        Dim objectTrxTime As String
        Dim customerID As String
        Dim customerName As String
        Dim featureID As String
        Dim featureName As String
        Dim featureUOM As String
        Dim featureAdultPrice As Decimal
        Dim featureChildPrice As Decimal
        Dim passbookID As String
        Dim passbookOwner As String
        Dim passbookOwnerRef As Customer
        Dim passbookVisitorName As String
        Dim passbookPurchDate As Date
        Dim passbookVisitorBDay As Date

        'Read the file
        Do While Not inputFile.EndOfStream
            line = inputFile.ReadLine
            If line.Length < 1 Then
                Continue Do
            End If

            If line.Chars(0) = "#" Then
                Continue Do
            End If

            field = Split(line, ";")
            objectTrxDate = Trim(field(0))
            objectTrxTime = Trim(field(1))
            objectName = Trim(field(2))
            objectAction = Trim(field(3))

            ' Dump the line to the transaction log
            txtTrxLogTabLog.Text &= vbCrLf _
                & "DISK FILE RECORD:" _
                & vbCrLf _
                & " - Record: Line=" & line
            For fieldIdx = 0 To field.Length - 1
                txtTrxLogTabLog.Text &=
                    vbCrLf &
                    "  >>> field " & fieldIdx.ToString & "='" & Trim(field(fieldIdx)) & "'"
            Next fieldIdx

            'Check if we have customer data
            If objectName = "CUSTOMER" Then
                customerID = Trim(field(4))
                customerName = Trim(field(5))
                If objectAction = "CREATE" Then
                    'See if Customer ID already exists
                    If mThemePark.findCustomer(customerID) = -1 Then
                        mThemePark.addCustomer(customerID, customerName)
                    Else
                        txtTrxLogTabLog.Text &= vbCrLf _
                            & "!! DUPLICATE FEATURE NOT ADDED !!" _
                            & vbCrLf
                    End If
                    Continue Do
                End If
            End If 'customer data

            'check if we have feature data
            If objectName = "FEATURE" Then
                featureID = Trim(field(4))
                featureName = Trim(field(5))
                featureUOM = Trim(field(6))
                featureAdultPrice = Decimal.Parse(Trim(field(7)))
                featureChildPrice = Decimal.Parse(Trim(field(8)))
                If objectAction = "CREATE" Then
                    'See if Feature ID already exists
                    If mThemePark.findFeature(featureID) = -1 Then
                        mThemePark.addFeature(featureID,
                                          featureName,
                                          featureUOM,
                                          featureAdultPrice,
                                          featureChildPrice)
                    Else
                        txtTrxLogTabLog.Text &= vbCrLf _
                            & "!! DUPLICATE FEATURE - NOT ADDED !!" _
                            & vbCrLf
                    End If
                End If
                Continue Do
            End If 'featureData

            'check if we have passbook data
            If objectName = "PASSBOOK" Then
                passbookID = Trim(field(4))
                passbookOwner = Trim(field(5))
                passbookPurchDate = Date.Parse(Trim(field(6)))
                passbookVisitorName = Trim(field(7))
                passbookVisitorBDay = Date.Parse(Trim(field(8)))
                'Get the customer reference
                idx = mThemePark.findCustomer(passbookOwner)
                If idx < 0 Then
                    'We can't get customer reference, something is wrong
                    'Give up in a blaze of glory.
                    txtTrxLogTabLog.Text &= vbCrLf _
                            & "!! PASSBOOK OWNER NOT FOUND !!" _
                            & vbCrLf
                    Continue Do
                End If
                passbookOwnerRef = mThemePark.ithCustomer(idx)

                'See if passbook ID already exists
                If mThemePark.findPassbook(passbookID) = -1 Then
                    mThemePark.addPassbook(
                        passbookID,
                        passbookOwnerRef,
                        passbookPurchDate,
                        passbookVisitorName,
                        passbookVisitorBDay)
                Else
                    txtTrxLogTabLog.Text &= vbCrLf _
                            & "!! DUPLICATE FEATURE - NOT ADDED !!" _
                            & vbCrLf
                End If

            End If 'passbook data

        Loop

        inputFile.Close()
    End Sub 'btnReadFile_Click


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

        'Update the Count
        txtNumUsedFeatureTabDashboard.Text = mThemePark.numUsedFeatures.ToString


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
