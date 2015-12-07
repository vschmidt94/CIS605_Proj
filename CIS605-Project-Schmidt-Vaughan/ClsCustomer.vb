'Template Copyright (c) 2009-2015 Dan Turk

#Region "Class / File Comment Header block"
'Program:               Theme Park Project          
'File:                  ClsCustomer    
'Author:                Vaughan Schmidt  
'Description:           Customer Class        
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

Public Class Customer

#Region "Attributes"
    '******************************************************************
    'Attributes + Module-level Constants+Variables
    '******************************************************************

    'No Attributes are currently defined.

    '********** Module-level constants

    '********** Module-level variables'
    Private mCustID As String
    Private mCustName As String

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
    ''' Specialty constructor, creates new Customer object with 
    ''' supplied ID and Name parameters
    ''' </summary>
    ''' <param name="pCustID">Customer ID</param>
    ''' <param name="pCustName">Customer Name</param>
    Public Sub New(ByVal pCustID As String,
                   ByVal pCustName As String)

        MyBase.New()
        _custID = pCustID
        _custName = pCustName

    End Sub 'New()

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

    Public Property custID() As String
        Get
            Return _custID
        End Get
        Set(ByVal pValue As String)
            _custID = pValue
        End Set
    End Property 'custID

    Public Property custName() As String
        Get
            Return _custName
        End Get
        Set(ByVal pValue As String)
            _custName = pValue
        End Set
    End Property 'custName

    '********** Private Get/Set Methods
    '             - access attributes, begin name with underscore (_)

    Private Property _custID() As String
        Get
            Return mCustID
        End Get
        Set(ByVal pValue As String)
            mCustID = pValue
        End Set
    End Property '_custID

    Private Property _custName As String
        Get
            Return mCustName
        End Get
        Set(ByVal pValue As String)
            mCustName = pValue
        End Set
    End Property

#End Region 'Get/Set Methods

#Region "Behavioral Methods"
    '******************************************************************
    'Behavioral Methods
    '******************************************************************

    'No Behavioral Methods are currently defined.

    '********** Public Shared Behavioral Methods

    ''' <summary>
    ''' Returns the string representation of the Customer object.
    ''' Overrides the inherited ToString() method.
    ''' </summary>
    ''' <returns>Customer object as String</returns>
    Public Overrides Function ToString() As String

        'Returns the value from the private function version
        Return _toString()

    End Function 'ToString()

    '********** Private Shared Behavioral Methods

    '********** Public Non-Shared Behavioral Methods

    '********** Private Non-Shared Behavioral Methods

    ''' <summary>
    ''' Returns a Customer object in String form.
    ''' </summary>
    ''' <returns>Customer object as String</returns>
    Private Function _toString() As String

        Dim tmpString As String
        tmpString = "( CUSTOMER: CustomerID=" & mCustID _
            & ", CustomerName=" & mCustName _
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

End Class 'ClsCustomer

