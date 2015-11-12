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
'Events:                The following events are defined:
'                           -ThemePark_CustomerAdded
'                           -ThemePark_FeatureAdded
'                           -ThemePark_PassbookAdded
'                           -ThemePark_PassbookFeatureAdded
'                           -ThemePark_UsedFeatureAdded
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
    Public Function addCustomer(ByVal pCustID As String,
                              ByVal pCustName As String) _
        As Customer

        ' Call the private function to do the work
        Return (_addCustomer(pCustID, pCustName))

    End Function 'addCustomer()

    ''' <summary>
    ''' Creates a new feature object, increases feature count
    ''' </summary>
    Public Function addFeature(ByVal pFeatureID As String,
                          ByVal pFeatureName As String,
                          ByVal pFeatureUOM As String,
                          ByVal pFeatureAdultPrice As Decimal,
                          ByVal pFeatureChildPrice As Decimal) _
        As Feature

        ' Call the private function to do the work
        Return (_addFeature(pFeatureID, pFeatureName, pFeatureUOM, pFeatureAdultPrice, pFeatureChildPrice))

    End Function 'addFeature()

    ''' <summary>
    ''' Creates a new Passbook object, increases passbook count
    ''' </summary>
    Public Function addPassbook(ByVal pPassbookID As String,
                           ByVal pPassbookOwner As Customer,
                           ByVal pPassbookDatePurch As Date,
                           ByVal pPassbookVisitorName As String,
                           ByVal pPassbookVisitorBirthdate As Date) _
        As Passbook

        ' Call the private function to do the work
        Return (_addPassbook(pPassbookID,
                             pPassbookOwner,
                             pPassbookDatePurch,
                             pPassbookVisitorName,
                             pPassbookVisitorBirthdate))

    End Function 'addPassbook()

    ''' <summary>
    ''' Creates a new PassbookFeature object, increases passbookFeature count
    ''' </summary>
    Public Function addPassbookFeature(ByVal pPassbookFeatureID As String,
                                  ByVal pQtyPurchased As Decimal,
                                  ByVal pPassbookFeatureAmt As Decimal,
                                  ByVal pPassbook As Passbook,
                                  ByVal pFeature As Feature,
                                  ByVal pQtyRemaining As Decimal) _
        As PassbookFeature

        ' Call the private function to do the work
        Return (_addPassbookFeature(pPassbookFeatureID,
                            pQtyPurchased,
                            pPassbookFeatureAmt,
                            pPassbook,
                            pFeature,
                            pQtyRemaining))

    End Function 'addPassbookFeature()

    ''' <summary>
    ''' Creates a new PassbookUsedFeature object, increases passbookFeature count
    ''' </summary>
    Public Function addUsedFeature(ByVal pUsedFeatureID As String,
                              ByVal pUsedPassbookFeature As PassbookFeature,
                              ByVal pUsedDate As Date,
                              ByVal pUsedLocation As String,
                              ByVal pQtyUsed As Decimal) _
        As UsedFeature

        ' Call the private function to do the work
        Return (_addUsedFeature(pUsedFeatureID,
                        pUsedPassbookFeature,
                        pUsedDate,
                        pUsedLocation,
                        pQtyUsed))

    End Function 'addUsedFeature()

    ''' <summary>
    ''' Updates a new PassbookFeature object, increases passbookFeature count
    ''' </summary>
    Public Function updatePassbookFeature(ByVal pPassbookFeature As PassbookFeature,
                                          ByVal pDateUpdated As Date,
                                          ByVal pUpdatedQty As Decimal) _
            As PassbookFeature

        ' Call the private function to do the work
        Return (_updatePassbookFeature(pPassbookFeature,
                            pDateUpdated,
                            pUpdatedQty))

    End Function 'updatePassbookFeature()

    '********** Private Non-Shared Behavioral Methods
    ''' <summary>
    ''' Creates a new customer object, increases customer count
    ''' </summary>
    Private Function _addCustomer(ByVal pCustID As String,
                                  ByVal pCustName As String) _
        As Customer

        ' Call the specialty constructor
        mNewCustomer = New Customer(pCustID, pCustName)

        ' Increase the customer count
        _numCustomers += 1

        ' Raise event
        RaiseEvent ThemePark_CustomerAdded(
            Me,
            New ThemePark_EventArgs_CustomerAdded(
                mNewCustomer
                )
            ) 'RaiseEvent

        Return mNewCustomer

    End Function '_addCustomer()

    ''' <summary>
    ''' Creates a new feature object, increases feature count
    ''' </summary>
    Private Function _addFeature(ByVal pFeatureID As String,
                                 ByVal pFeatureName As String,
                                 ByVal pFeatureUOM As String,
                                 ByVal pFeatureAdultPrice As Decimal,
                                 ByVal pFeatureChildPrice As Decimal) _
        As Feature

        ' Call the specialty constructor
        mNewFeature = New Feature(pFeatureID,
                                    pFeatureName,
                                    pFeatureUOM,
                                    pFeatureAdultPrice,
                                    pFeatureChildPrice)

        ' Increase the Feature count
        _numFeatures += 1

        ' Raise event
        RaiseEvent ThemePark_FeatureAdded(
            Me,
            New ThemePark_EventArgs_FeatureAdded(
                mNewFeature
                )
            ) 'RaiseEvent

        Return mNewFeature

    End Function '_addFeature()

    ''' <summary>
    ''' Creates a new Passbook object, increases passbook count
    ''' </summary>
    Private Function _addPassbook(ByVal pPassbookID As String,
                                  ByVal pPassbookOwner As Customer,
                                  ByVal pPassbookDatePurch As Date,
                                  ByVal pPassbookVisitorName As String,
                                  ByVal pPassbookVisitorBirthdate As Date) _
        As Passbook

        ' Call the specialty constructor
        mNewPassbook = New Passbook(pPassbookID,
                                    pPassbookOwner,
                                    pPassbookDatePurch,
                                    pPassbookVisitorName,
                                    pPassbookVisitorBirthdate)

        ' Increase the Passbook count
        _numPassbooks += 1

        ' Raise event
        RaiseEvent ThemePark_PassbookAdded(
            Me,
            New ThemePark_EventArgs_PassbookAdded(
                mNewPassbook
                )
            ) 'RaiseEvent

        Return mNewPassbook

    End Function '_addPassbook()

    ''' <summary>
    ''' Creates a new PassbookFeature object, increases passbookFeature count
    ''' </summary>
    Private Function _addPassbookFeature(ByVal pPassbookFeatureID As String,
                                         ByVal pQtyPurchased As Decimal,
                                         ByVal pPassbookFeatureAmt As Decimal,
                                         ByVal pPassbook As Passbook,
                                         ByVal pFeature As Feature,
                                         ByVal pQtyRemaining As Decimal) _
        As PassbookFeature

        ' Call the specialty constructor
        mNewPassbookFeature = New PassbookFeature(pPassbookFeatureID,
                                                  pQtyPurchased,
                                                  pPassbookFeatureAmt,
                                                  pPassbook,
                                                  pFeature,
                                                  pQtyRemaining)

        ' Increase the Feature count
        _numPassbookFeatures += 1

        ' Raise event
        RaiseEvent ThemePark_PassbookFeatureAdded(
            Me,
            New ThemePark_EventArgs_PassbookFeatureAdded(
                mNewPassbookFeature
                )
            ) 'RaiseEvent

        ' Return Ref to Object
        Return mNewPassbookFeature

    End Function '_addPassbookFeature()

    ''' <summary>
    ''' Creates a new UsedFeature object, increases passbookFeature count
    ''' </summary>
    Private Function _addUsedFeature(ByVal pUsedFeatureID As String,
                                     ByVal pUsedPassbookFeature As PassbookFeature,
                                     ByVal pUsedDate As Date,
                                     ByVal pUsedLocation As String,
                                     ByVal pQtyUsed As Decimal) _
        As UsedFeature

        ' Call the specialty constructor
        mNewUsedFeature = New UsedFeature(pUsedFeatureID,
                                          pUsedPassbookFeature,
                                          pUsedDate,
                                          pUsedLocation,
                                          pQtyUsed)

        ' Increase the Feature count
        _numUsedFeatures += 1

        ' Raise event
        RaiseEvent ThemePark_UsedFeatureAdded(
            Me,
            New ThemePark_EventArgs_UsedFeatureAdded(
                mNewUsedFeature
                )
            ) 'RaiseEvent

        Return mNewUsedFeature

    End Function '_addUsedFeature()

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

    ''' <summary>
    ''' Updates a new PassbookFeature object, increases passbookFeature count
    ''' </summary>
    Public Function _updatePassbookFeature(ByVal pPassbookFeature As PassbookFeature,
                                           ByVal pUpdatedDate As Date,
                                           ByVal pUpdatedQty As Decimal) _
        As PassbookFeature

        'update the feature in the selected passbook feature
        pPassbookFeature.passbookFeatureAmt = pPassbookFeature.passbookFeatureAmt + pUpdatedQty

        'Raise event
        RaiseEvent ThemePark_PassbookFeatureUpdated(
            Me,
            New ThemePark_EventArgs_PassbookFeatureUpdated(
                pPassbookFeature
                )
            ) 'Raise Event

        'return the modified passbook feature object
        Return (pPassbookFeature)

    End Function 'updatePassbookFeature()

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

    'These are all public.

    Public Event ThemePark_CustomerAdded(
        ByVal sender As System.Object,
        ByVal e As System.EventArgs
        )

    Public Event ThemePark_FeatureAdded(
        ByVal sender As System.Object,
        ByVal e As System.EventArgs
        )

    Public Event ThemePark_PassbookAdded(
        ByVal sender As System.Object,
        ByVal e As System.EventArgs
        )

    Public Event ThemePark_UsedFeatureAdded(
        ByVal sender As System.Object,
        ByVal e As System.EventArgs
        )

    Public Event ThemePark_PassbookFeatureAdded(
        ByVal sender As System.Object,
        ByVal e As System.EventArgs
        )

    Public Event ThemePark_PassbookFeatureUpdated(
        ByVal sender As System.Object,
        ByVal e As System.EventArgs
        )

#End Region 'Events

End Class 'ClsThemePark


