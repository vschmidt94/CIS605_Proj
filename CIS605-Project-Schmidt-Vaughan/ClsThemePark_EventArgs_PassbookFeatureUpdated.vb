'Template Copyright (c) 2009-2015 Dan Turk

#Region "Class / File Comment Header block"
'Program:               Theme Park Project
'File:                  ClsThemePark_EventArgs_FeatureUpdated.vb
'Author:                Vaughan Schmidt
'Description:           Event Args for Custom Feature Updated Events
'Date:                  November 8, 2015
'Tier:                  Business Logic
'Exceptions:         
'Exception-Handling: 
'Events:             
'Event-Handling:     
#End Region 'Class / File Comment Header block

#Region "Option / Imports"
Option Explicit On      'Must declare variables before using them
Option Strict On        'Must perform explicit data type conversions
#End Region 'Option / Imports

Public Class ThemePark_EventArgs_PassbookFeatureUpdated
    Inherits System.EventArgs

#Region "Attributes"
    '******************************************************************
    'Attributes + Module-level Constants+Variables
    '******************************************************************

    'No Attributes are currently defined.

    '********** Module-level constants

    '********** Module-level variables

    Private mTheUpdatedPassbookFeature As PassbookFeature

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
    ''' Crethe the EventArgs object
    ''' </summary>
    ''' <param name="pUpdatedPassbookFeature">the Passbook Feature that was updated</param>
    Public Sub New(
            ByVal pUpdatedPassbookFeature As PassbookFeature
            )

        MyBase.New
        _updatedPassbookFeature = pUpdatedPassbookFeature

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

    Public ReadOnly Property updatedPassbookFeature As PassbookFeature
        Get
            Return _updatedPassbookFeature
        End Get
    End Property 'usedFeature

    '********** Private Get/Set Methods
    '             - access attributes, begin name with underscore (_)

    Private Property _updatedPassbookFeature As PassbookFeature
        Get
            Return mTheUpdatedPassbookFeature
        End Get
        Set(pValue As PassbookFeature)
            mTheUpdatedPassbookFeature = pValue
        End Set
    End Property '_updatedPassbookFeature

#End Region 'Get/Set Methods

#Region "Behavioral Methods"
    '******************************************************************
    'Behavioral Methods
    '******************************************************************

    'No Behavioral Methods are currently defined.

    '********** Public Shared Behavioral Methods

    '********** Private Shared Behavioral Methods

    '********** Public Non-Shared Behavioral Methods

    Public Overrides Function ToString() As String

        Return _toString()

    End Function 'ToString()

    '********** Private Non-Shared Behavioral Methods

    Private Function _toString() As String

        Dim tmpStr As String
        tmpStr = "( THEMEPARK EVENT_ARGS PASSBOOK_FEATURE_UPDATED: " _
            & "ThePassbookFeature=" & _updatedPassbookFeature.ToString _
            & " )"

        Return tmpStr

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

End Class 'ThemePark_EventArgs_PassbookFeatureUpdated


