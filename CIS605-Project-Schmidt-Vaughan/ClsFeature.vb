'Template Copyright (c) 2009-2015 Dan Turk

#Region "Class / File Comment Header block"
'Program:               Theme Park Project          
'File:                  ClsFeature
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
Option Explicit On      'Must declare variables before using them
Option Strict On        'Must perform explicit data type conversions
#End Region 'Option / Imports

Public Class Feature

#Region "Attributes"
    '******************************************************************
    'Attributes + Module-level Constants+Variables
    '******************************************************************

    'No Attributes are currently defined.

    '********** Module-level constants
    Private mDEFAULT_FEATURE_ID As String = "FEAT001"
    Private mDEFAULT_FEATURE_NAME As String = "Default Feature"
    Private mDEFAULT_FEATURE_UOM As String = "HOUR"
    Private mDEFAULT_FEATURE_ADULT_PRICE As Decimal = 100D
    Private mDEFAULT_FEATURE_CHILD_PRICE As Decimal = 50D

    '********** Module-level variables
    Private mFeatureID As String
    Private mFeatureName As String
    Private mFeatureUOM As String
    Private mFeatureAdultPrice As Decimal
    Private mFeatureChildPrice As Decimal

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
    ''' Specialty constructor, creates new Feature Object with the supplied
    ''' parameters.
    ''' </summary>
    ''' <param name="pFeatureID">Feature ID</param>
    ''' <param name="pFeatureName">Feature Name</param>
    ''' <param name="pFeatureUOM">Feature UOM</param>
    ''' <param name="pFeatureAdultPrice">Feature Adult Price</param>
    ''' <param name="pFeatureChildPrice">Feature Child Price</param>
    Public Sub New(ByVal pFeatureID As String,
                    ByVal pFeatureName As String,
                    ByVal pFeatureUOM As String,
                    ByVal pFeatureAdultPrice As Decimal,
                    ByVal pFeatureChildPrice As Decimal)

        MyBase.New()
        _featureID = pFeatureID
        _featureName = pFeatureName
        _featureUOM = pFeatureUOM
        _featureAdultPrice = pFeatureAdultPrice
        _featureChildPrice = pFeatureChildPrice

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

    Public Property featureID() As String
        Get
            Return _featureID
        End Get
        Set(ByVal pValue As String)
            _featureID = pValue
        End Set
    End Property 'featureID

    Public Property featureName() As String
        Get
            Return _featureName
        End Get
        Set(ByVal pValue As String)
            _featureName = pValue
        End Set
    End Property 'featureName

    Public Property featureUOM() As String
        Get
            Return _featureUOM
        End Get
        Set(ByVal pValue As String)
            _featureUOM = pValue
        End Set
    End Property 'featureUOM

    Public Property featureAdultPrice() As Decimal
        Get
            Return _featureAdultPrice
        End Get
        Set(ByVal pValue As Decimal)
            _featureAdultPrice = pValue
        End Set
    End Property 'featureAdultPrice

    Public Property featureChildPrice() As Decimal
        Get
            Return _featureChildPrice
        End Get
        Set(ByVal pValue As Decimal)
            _featureChildPrice = pValue
        End Set
    End Property 'featureID

    '********** Private Get/Set Methods
    '             - access attributes, begin name with underscore (_)

    Private Property _featureID() As String
        Get
            Return mFeatureID
        End Get
        Set(ByVal pValue As String)
            mFeatureID = pValue
        End Set
    End Property '_FeatureID

    Private Property _featureName() As String
        Get
            Return mFeatureName
        End Get
        Set(ByVal pValue As String)
            mFeatureName = pValue
        End Set
    End Property '_FeatureName

    Public Property _featureUOM() As String
        Get
            Return mFeatureUOM
        End Get
        Set(ByVal pValue As String)
            mFeatureUOM = pValue
        End Set
    End Property '_featureUOM

    Public Property _featureAdultPrice() As Decimal
        Get
            Return mFeatureAdultPrice
        End Get
        Set(ByVal pValue As Decimal)
            mFeatureAdultPrice = pValue
        End Set
    End Property '_featureAdultPrice

    Public Property _featureChildPrice() As Decimal
        Get
            Return mFeatureChildPrice
        End Get
        Set(ByVal pValue As Decimal)
            mFeatureChildPrice = pValue
        End Set
    End Property '_featureID

#End Region 'Get/Set Methods

#Region "Behavioral Methods"
    '******************************************************************
    'Behavioral Methods
    '******************************************************************

    'No Behavioral Methods are currently defined.

    '********** Public Shared Behavioral Methods

    ''' <summary>
    ''' Returns the string representation of the Feature object.
    ''' Overrides the inherited ToString() method.
    ''' </summary>
    ''' <returns>Feature object as String</returns>
    Public Overrides Function ToString() As String

        'Returns the value from the private function version
        Return _toString()

    End Function 'ToString()

    '********** Private Shared Behavioral Methods

    '********** Public Non-Shared Behavioral Methods

    '********** Private Non-Shared Behavioral Methods

    ''' <summary>
    ''' Returns a Feature object in String form.
    ''' </summary>
    ''' <returns>Feature object as String</returns>
    Private Function _toString() As String

        Dim tmpString As String
        tmpString = "( FEATURE: FeatureID=" & mFeatureID _
            & ", FeatureName=" & mFeatureName _
            & ", FeatureUOM=" & mFeatureUOM _
            & ", FeatureAdultPrice=" & mFeatureAdultPrice.ToString _
            & ", FeatureChildPrice=" & mFeatureChildPrice.ToString _
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

End Class 'ClsFeature

