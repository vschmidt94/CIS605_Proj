'Template Copyright (c) 2009-2015 Dan Turk

#Region "Class / File Comment Header block"
'Program:               Theme Park Project          
'File:                  ClsPassbook         
'Author:                Vaughan Schmidt  
'Description:           Passbook Class        
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

Public Class Passbook

#Region "Attributes"
    '******************************************************************
    'Attributes + Module-level Constants+Variables
    '******************************************************************

    'No Attributes are currently defined.

    '********** Module-level constants

    Private Const mCHILD_AGE_CUTOFF As Integer = 13

    '********** Module-level variables

    Private mPassbookID As String
    Private mPassbookOwner As Customer
    Private mPassbookDatePurch As Date
    Private mPassbookVisitorName As String
    Private mPassbookVisitorBirthdate As Date

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

    Public Sub New(ByVal pPassbookID As String,
                    ByRef pPassbookOwner As Customer,
                    ByVal pPassbookDatePurch As Date,
                    ByVal pPassbookVisitorName As String,
                    ByVal pPassbookVisitorBirthdate As Date
                    )

        MyBase.New()
        _passbookID = pPassbookID
        _passbookOwner = pPassbookOwner
        _passbookDatePurch = pPassbookDatePurch
        _passbookVisitorName = pPassbookVisitorName
        _passbookVisitorBirthdate = pPassbookVisitorBirthdate

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

    Public Property passbookID() As String
        Get
            Return _passbookID
        End Get
        Set(ByVal pValue As String)
            _passbookID = pValue
        End Set
    End Property 'passbookID

    Public Property passbookOwner() As Customer
        Get
            Return _passbookOwner
        End Get
        Set(ByVal pValue As Customer)
            _passbookOwner = pValue
        End Set
    End Property 'passbookCustomer

    Public Property passbookDatePurch() As Date
        Get
            Return _passbookDatePurch
        End Get
        Set(ByVal pValue As Date)
            _passbookDatePurch = pValue
        End Set
    End Property 'passbookDatePurch

    Public Property passbookVisitorName() As String
        Get
            Return _passbookVisitorName
        End Get
        Set(ByVal pValue As String)
            _passbookVisitorName = pValue
        End Set
    End Property 'passbookVisitorName

    Public Property passbookVisitorBirthdate() As Date
        Get
            Return _passbookVisitorBirthdate
        End Get
        Set(ByVal pValue As Date)
            _passbookVisitorBirthdate = pValue
        End Set
    End Property 'passbookVisitorBirthdate

    Public ReadOnly Property age() As Integer
        Get
            Return _age
        End Get
    End Property 'age

    Public ReadOnly Property isChild() As Boolean
        Get
            Return _isChild
        End Get
    End Property 'isChild()

    '********** Private Get/Set Methods
    '             - access attributes, begin name with underscore (_)

    Private Property _passbookID() As String
        Get
            Return mPassbookID
        End Get
        Set(ByVal pValue As String)
            mPassbookID = pValue
        End Set
    End Property '_passbookID

    Private Property _passbookOwner() As Customer
        Get
            Return mPassbookOwner
        End Get
        Set(ByVal pValue As Customer)
            mPassbookOwner = pValue
        End Set
    End Property '_passbookCustomer

    Private Property _passbookDatePurch() As Date
        Get
            Return mPassbookDatePurch
        End Get
        Set(ByVal pValue As Date)
            mPassbookDatePurch = pValue
        End Set
    End Property '_passbookDatePurch

    Private Property _passbookVisitorName() As String
        Get
            Return mPassbookVisitorName
        End Get
        Set(ByVal pValue As String)
            mPassbookVisitorName = pValue
        End Set
    End Property '_passbookVisitorName

    Private Property _passbookVisitorBirthdate() As Date
        Get
            Return mPassbookVisitorBirthdate
        End Get
        Set(ByVal pValue As Date)
            mPassbookVisitorBirthdate = pValue
        End Set
    End Property '_passbookVisitorBirthdate

    Private ReadOnly Property _age() As Integer
        Get
            Dim numDays As Integer
            Dim Age As Integer

            numDays = CInt(DateDiff(DateInterval.Day, mPassbookVisitorBirthdate, Now))
            Age = CInt(numDays \ 365)
            Return Age
        End Get
    End Property '_age()

    Private ReadOnly Property _isChild() As Boolean
        Get
            'If the person is less than the cutoff age then they are child
            If (Me._age < mCHILD_AGE_CUTOFF) Then
                Return True
            End If
            Return False
        End Get
    End Property '_isChild()

#End Region 'Get/Set Methods

#Region "Behavioral Methods"
    '******************************************************************
    'Behavioral Methods
    '******************************************************************

    'No Behavioral Methods are currently defined.

    '********** Public Shared Behavioral Methodsa

    ''' <summary>
    ''' Returns the string representation of the Passbook object.
    ''' Overrides the inherited ToString() method.
    ''' </summary>
    ''' <returns>Passbook object as String</returns>
    Public Overrides Function ToString() As String

        'Returns the value from the private function version
        Return _toString()

    End Function 'ToString()

    '********** Private Shared Behavioral Methods

    '********** Public Non-Shared Behavioral Methods

    '********** Private Non-Shared Behavioral Methods

    ''' <summary>
    ''' Returns a Passbook object in String form.
    ''' </summary>
    ''' <returns>Passbook object as String</returns>
    Private Function _toString() As String

        Dim tmpString As String
        tmpString = "( Passbook: PassbookID=" & mPassbookID _
            & ", PassbookOwner=" & mPassbookOwner.ToString _
            & ", PassbookDatePurch=" & mPassbookDatePurch.ToString _
            & ", PassbookVisitorName=" & mPassbookVisitorName _
            & ", NumPassbookFeatures=" & mPassbookVisitorBirthdate.ToString _
            & ", VisitorAge=" & _age.ToString _
            & ", VisitorIsChild=" & _isChild.ToString _
            & " )"

        Return tmpString

    End Function '_toString(

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

End Class 'ClsPassbook

