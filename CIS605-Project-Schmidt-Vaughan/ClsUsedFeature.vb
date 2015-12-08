'Template Copyright (c) 2009-2015 Dan Turk

#Region "Class / File Comment Header block"
'Program:               Theme Park Project          
'File:                  ClsUsedFeature
'Author:                Vaughan Schmidt  
'Description:           Used Feature Class        
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

Public Class UsedFeature

#Region "Attributes"
    '******************************************************************
    'Attributes + Module-level Constants+Variables
    '******************************************************************

    'No Attributes are currently defined.

    '********** Module-level constants

    '********** Module-level variables
    Private mUsedFeatureID As String
    Private mUsedPassbookFeature As PassbookFeature
    Private mUsedDate As Date
    Private mUsedLocation As String
    Private mQtyUsed As Decimal

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
    ''' Specialty constructor for used passbook feature object
    ''' </summary>
    ''' <param name="pUsedFeatureID">Used Feature ID</param>
    ''' <param name="pUsedPassbookFeature">Used Passbook Feature</param>
    ''' <param name="pUsedDate">Date Used</param>
    ''' <param name="pUsedLocation">Location Used</param>
    ''' <param name="pQtyUsed">Qty Used</param>
    Public Sub New(ByVal pUsedFeatureID As String,
                    ByVal pUsedPassbookFeature As PassbookFeature,
                    ByVal pUsedDate As Date,
                    ByVal pUsedLocation As String,
                    ByVal pQtyUsed As Decimal)

        MyBase.New()
        _usedFeatureID = pUsedFeatureID
        _usedPassbookFeature = pUsedPassbookFeature
        _usedDate = pUsedDate
        _usedLocation = pUsedLocation
        _usedQty = pQtyUsed

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

    Public Property usedFeatureID() As String
        Get
            Return _usedFeatureID
        End Get
        Set(ByVal pValue As String)
            _usedFeatureID = pValue
        End Set
    End Property 'usedFeatureID()

    Public Property usedPassbookFeature() As PassbookFeature
        Get
            Return _usedPassbookFeature
        End Get
        Set(ByVal pValue As PassbookFeature)
            _usedPassbookFeature = pValue
        End Set
    End Property 'usedPassbookFeature()

    Public Property usedDate() As Date
        Get
            Return _usedDate
        End Get
        Set(ByVal pValue As Date)
            _usedDate = pValue
        End Set
    End Property 'usedDate()

    Public Property usedLocation() As String
        Get
            Return _usedLocation
        End Get
        Set(ByVal pValue As String)
            _usedLocation = pValue
        End Set
    End Property 'usedLocation()

    Public Property usedQty() As Decimal
        Get
            Return _usedQty
        End Get
        Set(ByVal pValue As Decimal)
            _usedQty = pValue
        End Set
    End Property 'usedQty()

    '********** Private Get/Set Methods
    '             - access attributes, begin name with underscore (_)

    Private Property _usedFeatureID() As String
        Get
            Return mUsedFeatureID
        End Get
        Set(ByVal pValue As String)
            mUsedFeatureID = pValue
        End Set
    End Property '_usedFeatureID()

    Private Property _usedPassbookFeature() As PassbookFeature
        Get
            Return mUsedPassbookFeature
        End Get
        Set(ByVal pValue As PassbookFeature)
            mUsedPassbookFeature = pValue
        End Set
    End Property '_usedPassbookFeature()

    Private Property _usedDate() As Date
        Get
            Return mUsedDate
        End Get
        Set(ByVal pValue As Date)
            mUsedDate = pValue
        End Set
    End Property '_usedDate()

    Private Property _usedLocation() As String
        Get
            Return mUsedLocation
        End Get
        Set(ByVal pValue As String)
            mUsedLocation = pValue
        End Set
    End Property '_usedLocation()

    Private Property _usedQty() As Decimal
        Get
            Return mQtyUsed
        End Get
        Set(ByVal pValue As Decimal)
            mQtyUsed = pValue
        End Set
    End Property '_QtyUsed()

#End Region 'Get/Set Methods

#Region "Behavioral Methods"
    '******************************************************************
    'Behavioral Methods
    '******************************************************************

    'No Behavioral Methods are currently defined.

    '********** Public Shared Behavioral Methods

    ''' <summary>
    ''' Returns the string representation of the UsedFeature object.
    ''' Overrides the inherited ToString() method.
    ''' </summary>
    ''' <returns>UsedFeature object as String</returns>
    Public Overrides Function ToString() As String

        'Returns the value from the private function version
        Return _toString()

    End Function 'ToString()


    '********** Private Shared Behavioral Methods

    '********** Public Non-Shared Behavioral Methods

    '********** Private Non-Shared Behavioral Methods

    ''' <summary>
    ''' Returns a UsedFeature object in String form.
    ''' </summary>
    ''' <returns>UsedFeature object as String</returns>
    Private Function _toString() As String

        Dim tmpString As String
        tmpString = "( USEDFEATURE: UsedFeatureID=" & mUsedFeatureID _
            & ", UsedPassbookFeature=" & mUsedPassbookFeature.ToString _
            & ", UsedFeatureDate=" & mUsedDate.ToString _
            & ", UsedFeatureLocation=" & mUsedLocation _
            & ", QtyUsed=" & mQtyUsed.ToString _
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

End Class 'ClsUsedFeature
