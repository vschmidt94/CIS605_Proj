Option Explicit On      'Must declare variables before using them
Option Strict On        'Must perform explicit data type conversions
'Template Copyright (c) 2009-2015 Dan Turk

#Region "Class / File Comment Header block"
'Program:               Theme Park Project          
'File:                  ClsThemePark
'Author:                Vaughan Schmidt  
'Description:           Theme Park Class        
'Date:                  2015.10.01
'Tier:                  Business Logic
'Exceptions:            TBD
'Exception-Handling:    TBD
'Events:                TBD        
'Event-Handling:        TBD
#End Region 'Class / File Comment Header block

#Region "Option / Imports"
Imports CIS605_Project_Schmidt_Vaughan
#End Region 'Option / Imports

Public Class ThemePark

#Region "Attributes"
    '******************************************************************
    'Attributes + Module-level Constants+Variables
    '******************************************************************

    'No Attributes are currently defined.

    '********** Module-level constants
    Private Const mDEFAULT_PARK_NAME As String = "Creataceous Park"
    Private Const mDEFAULT_NUM_CUST As Integer = 0
    Private Const mDEFAULT_NUM_PASSBOOKS As Integer = 0
    Private Const mDEFAULT_NUM_FEATURES As Integer = 0
    Private Const mDEFAULT_NUM_PASSBOOK_FEATURES As Integer = 0
    Private Const mDEFAULT_NUM_USED_FEATURES As Integer = 0

    '********** Module-level variables
    Private mParkName As String
    Private mNumCustomers As Integer
    Private mNumPassbooks As Integer
    Private mNumFeatures As Integer
    Private mNumPassbookFeatures As Integer
    Private mNumUsedFeatures As Integer
    Private mNewCustomer As Customer
    Private mNewFeature As Feature
    Private mNewPassbook As Passbook
    Public Property mNewPassbookFeature As PassbookFeature
    Public Property mNewUsedFeature As UsedFeature

#End Region 'Attributes

#Region "Constructors"
    '******************************************************************
    'Constructors
    '******************************************************************

    'No Constructors are currently defined.
    'These are all public.

    '********** Default constructor
    '             - no parameters

    ''' <summary>
    ''' Default constructor for ThemePark object
    ''' </summary>
    Public Sub New()

        MyBase.New()
        _parkName = mDEFAULT_PARK_NAME
        _numCustomers = mDEFAULT_NUM_CUST
        _numPassbooks = mDEFAULT_NUM_PASSBOOKS
        _numPassbookFeatures = mDEFAULT_NUM_PASSBOOK_FEATURES
        _numFeatures = mDEFAULT_NUM_FEATURES
        _numUsedFeatures = mDEFAULT_NUM_USED_FEATURES

    End Sub 'New() Default Constructor

    '********** Special constructor(s)
    '             - typically constructors have parameters 
    '               that are used to initialize attributes

    Public Sub New(ByVal pParkName As String)

        Me.New()
        _parkName = pParkName

    End Sub 'New() Specialty Constructor

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

    Public Property parkName() As String
        Get
            Return _parkName
        End Get
        Set(ByVal pValue As String)
            _parkName = pValue
        End Set
    End Property 'parkName()

    Public Property numCustomers() As Integer
        Get
            Return _numCustomers
        End Get
        Set(ByVal pValue As Integer)
            _numCustomers = pValue
        End Set
    End Property 'numCustomers

    Public Property numPassbooks() As Integer
        Get
            Return _numPassbooks
        End Get
        Set(ByVal pValue As Integer)
            _numPassbooks = pValue
        End Set
    End Property 'numPassbooks

    Public Property numPassbookFeatures() As Integer
        Get
            Return _numPassbookFeatures
        End Get
        Set(ByVal pValue As Integer)
            _numPassbookFeatures = pValue
        End Set
    End Property 'numPassbookFeatures

    Public Property numFeatures() As Integer
        Get
            Return _numFeatures
        End Get
        Set(ByVal pValue As Integer)
            _numFeatures = pValue
        End Set
    End Property 'numFeatures

    Public Property numUsedFeatures() As Integer
        Get
            Return _numUsedFeatures
        End Get
        Set(ByVal pValue As Integer)
            _numUsedFeatures = pValue
        End Set
    End Property 'numUsedFeatures

    '********** Private Get/Set Methods
    '             - access attributes, begin name with underscore (_)

    Private Property _parkName() As String
        Get
            Return mParkName
        End Get
        Set(ByVal pValue As String)
            mParkName = pValue
        End Set
    End Property '_parkName()

    Private Property _numCustomers() As Integer
        Get
            Return mNumCustomers
        End Get
        Set(ByVal pValue As Integer)
            mNumCustomers = pValue
        End Set
    End Property '_numCustomers

    Private Property _numPassbooks() As Integer
        Get
            Return mNumPassbooks
        End Get
        Set(ByVal pValue As Integer)
            mNumPassbooks = pValue
        End Set
    End Property '_numPassbooks

    Private Property _numPassbookFeatures() As Integer
        Get
            Return mNumPassbookFeatures
        End Get
        Set(ByVal pValue As Integer)
            mNumPassbookFeatures = pValue
        End Set
    End Property '_numPassbookFeatures

    Private Property _numFeatures() As Integer
        Get
            Return mNumFeatures
        End Get
        Set(ByVal pValue As Integer)
            mNumFeatures = pValue
        End Set
    End Property '_numFeatures

    Private Property _numUsedFeatures() As Integer
        Get
            Return mNumUsedFeatures
        End Get
        Set(ByVal pValue As Integer)
            mNumUsedFeatures = pValue
        End Set
    End Property '_numUsedFeatures



#End Region 'Get/Set Methods

#Region "Behavioral Methods"
    '******************************************************************
    'Behavioral Methods
    '******************************************************************

    'No Behavioral Methods are currently defined.

    '********** Public Shared Behavioral Methods

    ''' <summary>
    ''' Returns the string representation of the ThemePark object.
    ''' Overrides the inherited ToString() method.
    ''' </summary>
    ''' <returns>ThemePark object as String</returns>
    Public Overrides Function ToString() As String

        'Returns the value from the private function version
        Return _toString()

    End Function 'ToString()

    '********** Private Shared Behavioral Methods

    '********** Public Non-Shared Behavioral Methods

    ''' <summary>
    ''' Creates a new customer object, increases customer count
    ''' </summary>
    Public Sub addCustomer(ByVal pCustID As String,
                              ByVal pCustName As String)

        ' Call the private function to do the work
        _addCustomer(pCustID, pCustName)

    End Sub 'addCustomer()

    ''' <summary>
    ''' Creates a new feature object, increases feature count
    ''' </summary>
    Public Sub addFeature(ByVal pFeatureID As String,
                          ByVal pFeatureName As String,
                          ByVal pFeatureUOM As String,
                          ByVal pFeatureAdultPrice As Decimal,
                          ByVal pFeatureChildPrice As Decimal)

        ' Call the private function to do the work
        _addFeature(pFeatureID, pFeatureName, pFeatureUOM, pFeatureAdultPrice, pFeatureChildPrice)

    End Sub 'addFeature()

    ''' <summary>
    ''' Creates a new Passbook object, increases passbook count
    ''' </summary>
    Public Sub addPassbook(ByVal pPassbookID As String,
                           ByVal pPassbookOwner As Customer,
                           ByVal pPassbookDatePurch As Date,
                           ByVal pPassbookVisitorName As String,
                           ByVal pPassbookVisitorBirthdate As Date)

        ' Call the private function to do the work
        _addPassbook(pPassbookID, pPassbookOwner, pPassbookDatePurch, pPassbookVisitorName, pPassbookVisitorBirthdate)

    End Sub 'addPassbook()

    ''' <summary>
    ''' Creates a new PassbookFeature object, increases passbookFeature count
    ''' </summary>
    Public Sub addPassbookFeature(ByVal pPassbookFeatureID As String,
                                  ByVal pQtyPurchased As Decimal,
                                  ByVal pPassbookFeatureAmt As Decimal,
                                  ByVal pPassbook As Passbook,
                                  ByVal pFeature As Feature,
                                  ByVal pQtyRemaining As Decimal)

        ' Call the private function to do the work
        _addPassbookFeature(pPassbookFeatureID,
                            pQtyPurchased,
                            pPassbookFeatureAmt,
                            pPassbook,
                            pFeature,
                            pQtyRemaining)

    End Sub 'addPassbook()

    ''' <summary>
    ''' Creates a new PassbookUsedFeature object, increases passbookFeature count
    ''' </summary>
    Public Sub addUsedFeature(ByVal pUsedFeatureID As String,
                              ByVal pUsedPassbookFeature As PassbookFeature,
                              ByVal pUsedDate As Date,
                              ByVal pUsedLocation As String,
                              ByVal pQtyUsed As Decimal)

        ' Call the private function to do the work
        _addUsedFeature(pUsedFeatureID,
                        pUsedPassbookFeature,
                        pUsedDate,
                        pUsedLocation,
                        pQtyUsed)

    End Sub 'addPassbook()

    '********** Private Non-Shared Behavioral Methods
    ''' <summary>
    ''' Creates a new customer object, increases customer count
    ''' </summary>
    Private Sub _addCustomer(ByVal pCustID As String,
                             ByVal pCustName As String)

        ' Call the specialty constructor
        mNewCustomer = New Customer(pCustID, pCustName)

        ' Increase the customer count
        _numCustomers += 1

    End Sub '_addCustomer()

    ''' <summary>
    ''' Creates a new feature object, increases feature count
    ''' </summary>
    Private Sub _addFeature(ByVal pFeatureID As String,
                            ByVal pFeatureName As String,
                            ByVal pFeatureUOM As String,
                            ByVal pFeatureAdultPrice As Decimal,
                            ByVal pFeatureChildPrice As Decimal)

        ' Call the specialty constructor
        mNewFeature = New Feature(pFeatureID,
                                    pFeatureName,
                                    pFeatureUOM,
                                    pFeatureAdultPrice,
                                    pFeatureChildPrice)

        ' Increase the Feature count
        _numFeatures += 1

    End Sub '_addFeature()

    ''' <summary>
    ''' Creates a new Passbook object, increases passbook count
    ''' </summary>
    Private Sub _addPassbook(ByVal pPassbookID As String,
                             ByVal pPassbookOwner As Customer,
                             ByVal pPassbookDatePurch As Date,
                             ByVal pPassbookVisitorName As String,
                             ByVal pPassbookVisitorBirthdate As Date)

        ' Call the specialty constructor
        mNewPassbook = New Passbook(pPassbookID,
                                    pPassbookOwner,
                                    pPassbookDatePurch,
                                    pPassbookVisitorName,
                                    pPassbookVisitorBirthdate)

        ' Increase the Feature count
        _numPassbooks += 1

    End Sub '_addPassbook()

    ''' <summary>
    ''' Creates a new PassbookFeature object, increases passbookFeature count
    ''' </summary>
    Private Sub _addPassbookFeature(ByVal pPassbookFeatureID As String,
                                    ByVal pQtyPurchased As Decimal,
                                    ByVal pPassbookFeatureAmt As Decimal,
                                    ByVal pPassbook As Passbook,
                                    ByVal pFeature As Feature,
                                    ByVal pQtyRemaining As Decimal)

        ' Call the specialty constructor
        mNewPassbookFeature = New PassbookFeature(pPassbookFeatureID,
                                                  pQtyPurchased,
                                                  pPassbookFeatureAmt,
                                                  pPassbook,
                                                  pFeature,
                                                  pQtyRemaining)

        ' Increase the Feature count
        _numPassbookFeatures += 1

    End Sub '_addPassbookFeature()

    ''' <summary>
    ''' Creates a new UsedFeature object, increases passbookFeature count
    ''' </summary>
    Private Sub _addUsedFeature(ByVal pUsedFeatureID As String,
                                ByVal pUsedPassbookFeature As PassbookFeature,
                                ByVal pUsedDate As Date,
                                ByVal pUsedLocation As String,
                                ByVal pQtyUsed As Decimal)

        ' Call the specialty constructor
        mNewUsedFeature = New UsedFeature(pUsedFeatureID,
                                          pUsedPassbookFeature,
                                          pUsedDate,
                                          pUsedLocation,
                                          pQtyUsed)

        ' Increase the Feature count
        _numUsedFeatures += 1

    End Sub '_addUsedFeature()

    ''' <summary>
    ''' Returns a Theme Park object in String form.
    ''' </summary>
    ''' <returns>Theme Park object as String</returns>
    Private Function _toString() As String

        Dim tmpString As String
        tmpString = "( THEMEPARK: ParkName=" & mParkName _
            & ", NumCustomers=" & mNumCustomers.ToString _
            & ", NumPassbooks=" & mNumPassbooks.ToString _
            & ", NumFeatures=" & mNumFeatures.ToString _
            & ", NumPassbookFeatures=" & mNumPassbookFeatures.ToString _
            & ", NumUsedFeatures=" & mNumUsedFeatures.ToString _
            & " )"

        Return tmpString

    End Function '_toString()

#End Region 'Behavioral Methods

#Region "Event Procedures"
    '******************************************************************
    'Event Procedures
    '******************************************************************

    'No Event Procedures are currently defined.
    'These are all private.

    '********** User-Interface Event Procedures
    '             - Initiated explicitly by user

    '********** User-Interface Event Procedures
    '             - Initiated automatically by system

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

End Class 'ClsThemePark

