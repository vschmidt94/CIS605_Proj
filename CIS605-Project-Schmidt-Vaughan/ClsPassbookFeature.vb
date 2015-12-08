'Template Copyright (c) 2009-2015 Dan Turk

#Region "Class / File Comment Header block"
'Program:               Theme Park Project          
'File:                  ClsPassbookFeature
'Author:                Vaughan Schmidt  
'Description:           Passbook Feature Class        
'Date:                  2015.10.01
'Tier:                  Business Logic
'Exceptions:            TBD
'Exception-Handling:    TBD
'Events:                TBD        
'Event-Handling:        TBD
#End Region 'Class / File Comment Header block

#Region "Option / Imports"
Option Explicit On      'Must declare variables before using them
Option Strict On        'Must perform explicit data type conversions
#End Region 'Option / Imports

Public Class PassbookFeature

#Region "Attributes"
    '******************************************************************
    'Attributes + Module-level Constants+Variables
    '******************************************************************

    'No Attributes are currently defined.

    '********** Module-level constants

    '********** Module-level variables
    Private mPassbookFeatureID As String
    Private mPassbookFeatureQtyPurch As Decimal
    Private mPassbookFeatureAmt As Decimal
    Private mPassbook As Passbook
    Private mFeature As Feature
    Private mPassbookFeatureQtyRemaining As Decimal
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

    ''' <summary>
    ''' Specialty constructore for PassbookFeature object
    ''' </summary>
    ''' <param name="pPassbookFeatureID">Passbook Feature ID</param>
    ''' <param name="pQtyPurchased">Qty Purchased</param>
    ''' <param name="pPassbookFeatureAmt">Feature Amount</param>
    ''' <param name="pPassbook">Passbook</param>
    ''' <param name="pFeature">Feature</param>
    ''' <param name="pQtyRemaining">Qty Remaining</param>
    Public Sub New(ByVal pPassbookFeatureID As String,
                    ByVal pQtyPurchased As Decimal,
                    ByVal pPassbookFeatureAmt As Decimal,
                    ByVal pPassbook As Passbook,
                    ByVal pFeature As Feature,
                    ByVal pQtyRemaining As Decimal)

        MyBase.New()
        _passbookFeatureID = pPassbookFeatureID
        _passbookFeatureQtyPurch = pQtyPurchased
        _passbookFeatureAmt = pPassbookFeatureAmt
        _passbook = pPassbook
        _feature = pFeature
        _passbookFeatureQtyRemaining = pQtyPurchased

    End Sub 'New() specialty constructor

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

    Public Property passbookFeatureID() As String
        Get
            Return _passbookFeatureID
        End Get
        Set(ByVal pValue As String)
            _passbookFeatureID = pValue
        End Set
    End Property 'passbookFeatureID

    Public Property passbookFeatureQtyPurch() As Decimal
        Get
            Return _passbookFeatureQtyPurch
        End Get
        Set(ByVal pValue As Decimal)
            _passbookFeatureQtyPurch = pValue
        End Set
    End Property 'passbookFeatureQtyPurch

    Public Property passbookFeatureAmt() As Decimal
        Get
            Return _passbookFeatureAmt
        End Get
        Set(ByVal pValue As Decimal)
            _passbookFeatureAmt = pValue
        End Set
    End Property 'passbookFeatureQtyPurch

    Public Property passbook() As Passbook
        Get
            Return _passbook
        End Get
        Set(ByVal pValue As Passbook)
            _passbook = pValue
        End Set
    End Property 'passbook()

    Public Property feature() As Feature
        Get
            Return _feature
        End Get
        Set(ByVal pValue As Feature)
            _feature = pValue
        End Set
    End Property 'feature()

    Public Property passbookFeatureQtyRemaining() As Decimal
        Get
            Return _passbookFeatureQtyRemaining
        End Get
        Set(ByVal pValue As Decimal)
            _passbookFeatureQtyRemaining = pValue
        End Set
    End Property 'feature()

    '********** Private Get/Set Methods
    '             - access attributes, begin name with underscore (_)

    Private Property _passbookFeatureID() As String
        Get
            Return mPassbookFeatureID
        End Get
        Set(ByVal pValue As String)
            mPassbookFeatureID = pValue
        End Set
    End Property '_passbookFeatureID

    Private Property _passbookFeatureQtyPurch() As Decimal
        Get
            Return mPassbookFeatureQtyPurch
        End Get
        Set(ByVal pValue As Decimal)
            mPassbookFeatureQtyPurch = pValue
        End Set
    End Property '_passbookFeatureQtyPurch()

    Private Property _passbookFeatureAmt() As Decimal
        Get
            Return mPassbookFeatureAmt
        End Get
        Set(ByVal pValue As Decimal)
            mPassbookFeatureAmt = pValue
        End Set
    End Property '_passbookFeatureAmt()

    Private Property _passbook() As Passbook
        Get
            Return mPassbook
        End Get
        Set(ByVal pValue As Passbook)
            mPassbook = pValue
        End Set
    End Property '_passbook()

    Private Property _feature() As Feature
        Get
            Return mFeature
        End Get
        Set(ByVal pValue As Feature)
            mFeature = pValue
        End Set
    End Property '_passbook()

    Private Property _passbookFeatureQtyRemaining As Decimal
        Get
            Return mPassbookFeatureQtyRemaining
        End Get
        Set(pValue As Decimal)
            mPassbookFeatureQtyRemaining = pValue
        End Set
    End Property '_passbookFeatureQtyRemaining()

#End Region 'Get/Set Methods

#Region "Behavioral Methods"
    '******************************************************************
    'Behavioral Methods
    '******************************************************************

    'No Behavioral Methods are currently defined.

    '********** Public Shared Behavioral Methods
    ''' <summary>
    ''' Returns the string representation of the PassbookFeature object.
    ''' Overrides the inherited ToString() method.
    ''' </summary>
    ''' <returns>PassbookFeature object as String</returns>
    Public Overrides Function ToString() As String

        'Returns the value from the private function version
        Return _toString()

    End Function 'ToString()

    '********** Private Shared Behavioral Methods

    '********** Public Non-Shared Behavioral Methods

    '********** Private Non-Shared Behavioral Methods

    ''' <summary>
    ''' Returns a PassbookFeature object in String form.
    ''' </summary>
    ''' <returns>PassbookFeature object as String</returns>
    Private Function _toString() As String

        Dim tmpString As String
        tmpString = "( PASSBOOKFEATURE: PassbookFeatureID=" & mPassbookFeatureID _
            & ", PassbookFeatureQtyPurchased=" & mPassbookFeatureQtyPurch.ToString _
            & ", PassbookFeatureAmt=" & mPassbookFeatureAmt.ToString _
            & ", Passbook=" & mPassbook.ToString _
            & ", Feature=" & mFeature.ToString _
            & ", PassbookFeatureQtyRemaining=" & mPassbookFeatureQtyRemaining.ToString _
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

End Class 'ClsPassbookFeature
