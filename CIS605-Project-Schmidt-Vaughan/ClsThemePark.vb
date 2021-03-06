﻿Option Explicit On      'Must declare variables before using them
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

    'How to handle arrays.  Normally these would be bigger values, keep 
    'them small for now to allow more debugging.
    'I understand that each array could have its own values, and 
    'perhaps have a need to grow at different rates, but I'm more interested
    'in forcing growth cycles & ease of debugging for now, so using a 
    ' 'global' constant for this to apply to all arrays.
    Private Const mDEFAULT_ARRAY_SIZE As Integer = 3
    Private Const mDEFAULT_ARRAY_GROWTH_INCREMENT As Integer = 3

    '********** Module-level variables
    Private mParkName As String

    'Array tracking
    Private mNumCustomers As Integer
    Private mMaxCustomers As Integer
    Private mNumPassbooks As Integer
    Private mMaxPassbooks As Integer
    Private mNumFeatures As Integer
    Private mMaxFeatures As Integer
    Private mNumPassbookFeatures As Integer
    Private mMaxPassbookFeatures As Integer
    Private mNumUsedFeatures As Integer
    Private mMaxUsedFeatures As Integer

    'Objects Arrays
    Private mCustomers() As Customer
    Private mFeatures() As Feature
    Private mPassbooks() As Passbook
    Private mPassbookFeatures() As PassbookFeature
    Private mUsedFeatures() As UsedFeature

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

        'Initialize object arrays
        _maxCustomers = mDEFAULT_ARRAY_SIZE
        _maxFeatures = mDEFAULT_ARRAY_SIZE
        _maxPassbooks = mDEFAULT_ARRAY_SIZE
        _maxPassbookFeatures = mDEFAULT_ARRAY_SIZE
        _maxUsedFeatures = mDEFAULT_ARRAY_SIZE

        _numCustomers = mDEFAULT_NUM_CUST
        _numPassbooks = mDEFAULT_NUM_PASSBOOKS
        _numPassbookFeatures = mDEFAULT_NUM_PASSBOOK_FEATURES
        _numFeatures = mDEFAULT_NUM_FEATURES
        _numUsedFeatures = mDEFAULT_NUM_USED_FEATURES

        'ReDim the arrays to current Max
        ReDim mCustomers(_maxCustomers - 1)
        ReDim mFeatures(_maxFeatures - 1)
        ReDim mPassbookFeatures(_maxPassbookFeatures - 1)
        ReDim mPassbooks(_maxPassbooks - 1)
        ReDim mUsedFeatures(_maxUsedFeatures - 1)

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

    Public ReadOnly Property maxCustomers() As Integer
        Get
            Return _maxCustomers
        End Get
    End Property 'maxCustomers

    Public Property numPassbooks() As Integer
        Get
            Return _numPassbooks
        End Get
        Set(ByVal pValue As Integer)
            _numPassbooks = pValue
        End Set
    End Property 'numPassbooks

    Public ReadOnly Property maxPassbooks() As Integer
        Get
            Return _maxPassbooks
        End Get
    End Property 'maxPassbooks

    Public Property numPassbookFeatures() As Integer
        Get
            Return _numPassbookFeatures
        End Get
        Set(ByVal pValue As Integer)
            _numPassbookFeatures = pValue
        End Set
    End Property 'numPassbookFeatures

    Public ReadOnly Property maxPassbookFeatures() As Integer
        Get
            Return _maxPassbookFeatures
        End Get
    End Property 'maxPassbooksFeatures

    Public Property numFeatures() As Integer
        Get
            Return _numFeatures
        End Get
        Set(ByVal pValue As Integer)
            _numFeatures = pValue
        End Set
    End Property 'numFeatures

    Public ReadOnly Property maxFeatures() As Integer
        Get
            Return _maxFeatures
        End Get
    End Property 'maxFeatures

    Public Property numUsedFeatures() As Integer
        Get
            Return _numUsedFeatures
        End Get
        Set(ByVal pValue As Integer)
            _numUsedFeatures = pValue
        End Set
    End Property 'numUsedFeatures

    Public ReadOnly Property maxUsedFeatures() As Integer
        Get
            Return _maxUsedFeatures
        End Get
    End Property 'maxUsedFeatures

    'All the ith* properties start here

    Public ReadOnly Property ithCustomer(ByVal pN As Integer) As Customer
        Get
            Return _ithCustomer(pN)
        End Get
    End Property 'ithCustomer()

    Public ReadOnly Property ithPassbook(ByVal pN As Integer) As Passbook
        Get
            Return _ithPassbook(pN)
        End Get
    End Property 'ithPassbook()

    Public ReadOnly Property ithFeature(ByVal pN As Integer) As Feature
        Get
            Return _ithFeature(pN)
        End Get
    End Property 'ithFeature()

    Public ReadOnly Property ithPassbookFeature(ByVal pN As Integer) As PassbookFeature
        Get
            Return _ithPassbookFeature(pN)
        End Get
    End Property 'ithPassbookFeature()

    Public ReadOnly Property ithUsedFeature(ByVal pN As Integer) As UsedFeature
        Get
            Return _ithUsedFeature(pN)
        End Get
    End Property 'ithUsedFeature()

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

    Private Property _maxCustomers() As Integer
        Get
            Return mMaxCustomers
        End Get
        Set(ByVal pValue As Integer)
            mMaxCustomers = pValue
        End Set
    End Property '_maxCustomers

    Private Property _numPassbooks() As Integer
        Get
            Return mNumPassbooks
        End Get
        Set(ByVal pValue As Integer)
            mNumPassbooks = pValue
        End Set
    End Property '_numPassbooks

    Private Property _maxPassbooks() As Integer
        Get
            Return mMaxPassbooks
        End Get
        Set(ByVal pValue As Integer)
            mMaxPassbooks = pValue
        End Set
    End Property '_maxPassbooks

    Private Property _numPassbookFeatures() As Integer
        Get
            Return mNumPassbookFeatures
        End Get
        Set(ByVal pValue As Integer)
            mNumPassbookFeatures = pValue
        End Set
    End Property '_numPassbookFeatures

    Private Property _maxPassbookFeatures() As Integer
        Get
            Return mMaxPassbookFeatures
        End Get
        Set(ByVal pValue As Integer)
            mMaxPassbookFeatures = pValue
        End Set
    End Property '_maxPassbookFeatures

    Private Property _numFeatures() As Integer
        Get
            Return mNumFeatures
        End Get
        Set(ByVal pValue As Integer)
            mNumFeatures = pValue
        End Set
    End Property '_numFeatures

    Private Property _maxFeatures() As Integer
        Get
            Return mMaxFeatures
        End Get
        Set(ByVal pValue As Integer)
            mMaxFeatures = pValue
        End Set
    End Property '_maxFeatures

    Private Property _numUsedFeatures() As Integer
        Get
            Return mNumUsedFeatures
        End Get
        Set(ByVal pValue As Integer)
            mNumUsedFeatures = pValue
        End Set
    End Property '_numUsedFeatures

    Private Property _maxUsedFeatures() As Integer
        Get
            Return mMaxUsedFeatures
        End Get
        Set(ByVal pValue As Integer)
            mMaxUsedFeatures = pValue
        End Set
    End Property '_maxUsedFeatures

    ' All the _ith_ methods start here.

    ''' <summary>
    ''' Returns the customer at index pN in the customer array
    ''' </summary>
    ''' <param name="pN">the index of the customer to be returned</param>
    ''' <returns></returns>
    Private Property _ithCustomer(ByVal pN As Integer) As Customer
        Get
            If pN >= 0 And pN < _maxCustomers Then
                Return mCustomers(pN)
            Else
                Throw New IndexOutOfRangeException
            End If
        End Get
        Set(pValue As Customer)
            If pN >= 0 And pN < _maxCustomers Then
                mCustomers(pN) = pValue
            Else
                Throw New IndexOutOfRangeException
            End If
        End Set
    End Property '_ithCustomer

    ''' <summary>
    ''' Returns the feature at index pN in the feature array
    ''' </summary>
    ''' <param name="pN">the index of the feature to be returned</param>
    ''' <returns></returns>
    Private Property _ithFeature(ByVal pN As Integer) As Feature
        Get
            If pN >= 0 And pN < _maxFeatures Then
                Return mFeatures(pN)
            Else
                Throw New IndexOutOfRangeException
            End If
        End Get
        Set(pValue As Feature)
            If pN >= 0 And pN < _maxFeatures Then
                mFeatures(pN) = pValue
            Else
                Throw New IndexOutOfRangeException
            End If
        End Set
    End Property '_ithFeature

    ''' <summary>
    ''' Returns the passbook at index pN in the passbook feature array
    ''' </summary>
    ''' <param name="pN">the index of the passbook to be returned</param>
    ''' <returns></returns>
    Private Property _ithPassbook(ByVal pN As Integer) As Passbook
        Get
            If pN >= 0 And pN < _maxPassbooks Then
                Return mPassbooks(pN)
            Else
                Throw New IndexOutOfRangeException
            End If
        End Get
        Set(pValue As Passbook)
            If pN >= 0 And pN < _maxPassbooks Then
                mPassbooks(pN) = pValue
            Else
                Throw New IndexOutOfRangeException
            End If
        End Set
    End Property '_ithPassbook

    ''' <summary>
    ''' Returns the passbook feature at index pN in the passbook feature array
    ''' </summary>
    ''' <param name="pN">the index of the passbook feature to be returned</param>
    ''' <returns></returns>
    Private Property _ithPassbookFeature(ByVal pN As Integer) As PassbookFeature
        Get
            If pN >= 0 And pN < _maxPassbookFeatures Then
                Return mPassbookFeatures(pN)
            Else
                Throw New IndexOutOfRangeException
            End If
        End Get
        Set(pValue As PassbookFeature)
            If pN >= 0 And pN < _maxPassbookFeatures Then
                mPassbookFeatures(pN) = pValue
            Else
                Throw New IndexOutOfRangeException
            End If
        End Set
    End Property '_ithPassbookFeature

    ''' <summary>
    ''' Returns the used feature at index pN in the used feature array
    ''' </summary>
    ''' <param name="pN">the index of the used feature to be returned</param>
    ''' <returns></returns>
    Private Property _ithUsedFeature(ByVal pN As Integer) As UsedFeature
        Get
            If pN >= 0 And pN < _maxUsedFeatures Then
                Return mUsedFeatures(pN)
            Else
                Throw New IndexOutOfRangeException
            End If
        End Get
        Set(pValue As UsedFeature)
            If pN >= 0 And pN < _maxUsedFeatures Then
                mUsedFeatures(pN) = pValue
            Else
                Throw New IndexOutOfRangeException
            End If
        End Set
    End Property '_ithUsedFeature

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
    ''' Returns index number of customer based on ID, -1 if it does not exist in arrays
    ''' </summary>
    ''' <param name="pCustID"></param>
    ''' <returns></returns>
    Public Function findCustomer(ByVal pCustID As String) _
        As Integer

        'Call the private function to do the work
        Return (_findCustomer(pCustID))
    End Function 'findCustomer()

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
    ''' Returns index number of feature based on ID, -1 if it does not exist in arrays
    ''' </summary>
    ''' <param name="pFeatureID"></param>
    ''' <returns></returns>
    Public Function findFeature(ByVal pFeatureID As String) _
        As Integer

        'Call the private function to do the work
        Return (_findFeature(pFeatureID))
    End Function 'findFeature()

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
    ''' Returns index number of passbook based on ID, -1 if it does not exist in arrays
    ''' </summary>
    ''' <param name="pPassbookID"></param>
    ''' <returns></returns>
    Public Function findPassbook(ByVal pPassbookID As String) _
        As Integer

        'Call the private function to do the work
        Return (_findPassbook(pPassbookID))
    End Function 'findPassbook()

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
    ''' Returns index number of passbook feature based on ID, -1 if it does not exist in arrays
    ''' </summary>
    ''' <param name="pPassbookFeatureID"></param>
    ''' <returns></returns>
    Public Function findPassbookFeature(ByVal pPassbookFeatureID As String) _
        As Integer

        'Call the private function to do the work
        Return (_findPassbookFeature(pPassbookFeatureID))
    End Function 'findPassbookFeature()


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
    ''' Returns index number of used passbook feature based on ID, -1 if it does not exist in arrays
    ''' </summary>
    ''' <param name="pUsedPassbookFeatureID"></param>
    ''' <returns></returns>
    Public Function findUsedPassbookFeature(ByVal pUsedPassbookFeatureID As String) _
        As Integer

        'Call the private function to do the work
        Return (_findUsedPassbookFeature(pUsedPassbookFeatureID))
    End Function 'findUsedPassbookFeature()


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

    ''' <summary>
    ''' Calculates the top feature in the park based on occurances
    ''' </summary>
    ''' <returns>String representation of top feature</returns>
    Public Function calcTopFeature() _
        As String

        'Return the results of the private function
        Return _calcTopFeature()

    End Function 'calcTopFeature

    ''' <summary>
    ''' Calculates the number of vistors with birthdays today
    ''' </summary>
    ''' <returns></returns>
    Public Function calcBirthdaysInMonth() As Integer
        'Return the results of private function
        Return _calcBirthdaysInMonth()
    End Function 'calcBirthdaysInMonth()

    ''' <summary>
    ''' Returns the average age of passbook visitors
    ''' </summary>
    ''' <returns></returns>
    Public Function calcAvgAge() As Integer
        'Return the results of the private function
        Return _calcAvgAge()
    End Function 'calcAvgAge()

    ''' <summary>
    ''' Returns the average number of passbooks per customer
    ''' </summary>
    ''' <returns></returns>
    Public Function calcAvgPBperCust() As Double
        'Return the results of the private function
        Return _calcAvgPBperCust()
    End Function 'calcAvgPBperCust()


    ''' <summary>
    ''' Returns the sum of the unused passbook features in dollars
    ''' </summary>
    ''' <returns></returns>
    Public Function calcSumUnusedPBF() As Double
        'Return the results of the private function
        Return _calcSumUnusedPBF()
    End Function 'calcSumUnusedPBF()


    ''' <summary>
    ''' Returns the sum of the average unused feature balance per passbook
    ''' </summary>
    ''' <returns></returns>
    Public Function calcAvgUnusedPBFBal() As Double
        'Return the results of the private function
        Return _calcAvgUnusedPBFBal()
    End Function 'calcAvgUnusedPBFBal()

    ''' <summary>
    ''' Returns the ratio of features used / features purchased (dollar ratio)
    ''' </summary>
    ''' <returns></returns>
    Public Function calcPercentFeatUse() As Double
        'Return the results of the private function
        Return _calcPercentFeatUse()
    End Function 'calcPercentFeatUse()

    '********** Private Non-Shared Behavioral Methods

    ''' <summary>
    ''' Returns the index number of the customer with the specific cusstomer ID
    ''' Returns -1 if not found
    ''' </summary>
    ''' <param name="pCustID"></param>
    ''' <returns></returns>
    Private Function _findCustomer(ByVal pCustID As String) _
        As Integer

        'Declare variables
        Dim i As Integer = 0

        'Loop through existing customers and see if we have a match

        For i = 0 To _numCustomers - 1
            If ithCustomer(i).custID = pCustID Then
                Return i
            End If
        Next

        ' if nothing found, return -1
        Return -1

    End Function '_findCustomer


    ''' <summary>
    ''' Creates a new customer object, increases customer count
    ''' </summary>
    Private Function _addCustomer(ByVal pCustID As String,
                                  ByVal pCustName As String) _
        As Customer

        'Declare Variable
        Dim newCustomer As Customer

        ' Call the specialty constructor
        newCustomer = New Customer(pCustID, pCustName)

        'Check that the array is large enough for a new customer,
        'if not, grow the array.
        If _numCustomers >= _maxCustomers Then
            _maxCustomers += mDEFAULT_ARRAY_GROWTH_INCREMENT
            ReDim Preserve mCustomers(_maxCustomers - 1)
        End If

        ' Add the customer to the array in the correct index
        Try
            _ithCustomer(_numCustomers) = newCustomer
        Catch ex As Exception
            Throw New IndexOutOfRangeException
        End Try

        ' Increase the customer count
        _numCustomers += 1

        ' Raise event
        RaiseEvent ThemePark_CustomerAdded(
            Me,
            New ThemePark_EventArgs_CustomerAdded(
                newCustomer
                )
            ) 'RaiseEvent

        Return newCustomer

    End Function '_addCustomer()



    ''' <summary>
    ''' Returns the index number of the feature with the specific feature ID
    ''' Returns -1 if not found
    ''' </summary>
    ''' <param name="pFeatureID"></param>
    ''' <returns></returns>
    Private Function _findFeature(ByVal pFeatureID As String) _
        As Integer

        'Declare variables
        Dim i As Integer = 0

        'Loop through existing features and see if we have a match

        For i = 0 To _numFeatures - 1
            If ithFeature(i).featureID = pFeatureID Then
                Return i
            End If
        Next

        ' if nothing found, return -1
        Return -1

    End Function '_findFeature


    ''' <summary>
    ''' Creates a new feature object, increases feature count
    ''' </summary>
    Private Function _addFeature(ByVal pFeatureID As String,
                                 ByVal pFeatureName As String,
                                 ByVal pFeatureUOM As String,
                                 ByVal pFeatureAdultPrice As Decimal,
                                 ByVal pFeatureChildPrice As Decimal) _
        As Feature

        'Declare variables
        Dim newFeature As Feature

        ' Call the specialty constructor
        newFeature = New Feature(pFeatureID,
                                    pFeatureName,
                                    pFeatureUOM,
                                    pFeatureAdultPrice,
                                    pFeatureChildPrice)

        'Check that the array is large enough for a new feature,
        'if not, grow the array.
        If _numFeatures >= _maxFeatures Then
            _maxFeatures += mDEFAULT_ARRAY_GROWTH_INCREMENT
            ReDim Preserve mFeatures(_maxFeatures - 1)
        End If

        ' Add the feature to the array in the correct index
        Try
            _ithFeature(_numFeatures) = newFeature
        Catch ex As Exception
            Throw New IndexOutOfRangeException
        End Try

        ' Increase the Feature count
        _numFeatures += 1

        ' Raise event
        RaiseEvent ThemePark_FeatureAdded(
            Me,
            New ThemePark_EventArgs_FeatureAdded(
                newFeature
                )
            ) 'RaiseEvent

        Return newFeature

    End Function '_addFeature()


    ''' <summary>
    ''' Returns the index number of the passbook with the specific passbook ID
    ''' Returns -1 if not found
    ''' </summary>
    ''' <param name="pPassbookID"></param>
    ''' <returns></returns>
    Private Function _findPassbook(ByVal pPassbookID As String) _
        As Integer

        'Declare variables
        Dim i As Integer = 0

        'Loop through existing Passbooks and see if we have a match

        For i = 0 To _numPassbooks - 1
            If ithPassbook(i).passbookID = pPassbookID Then
                Return i
            End If
        Next

        ' if nothing found, return -1
        Return -1

    End Function '_findPassbook

    ''' <summary>
    ''' Creates a new Passbook object, increases passbook count
    ''' </summary>
    Private Function _addPassbook(ByVal pPassbookID As String,
                                  ByVal pPassbookOwner As Customer,
                                  ByVal pPassbookDatePurch As Date,
                                  ByVal pPassbookVisitorName As String,
                                  ByVal pPassbookVisitorBirthdate As Date) _
        As Passbook

        'Declare Variables
        Dim newPassbook As Passbook

        ' Call the specialty constructor
        newPassbook = New Passbook(pPassbookID,
                                    pPassbookOwner,
                                    pPassbookDatePurch,
                                    pPassbookVisitorName,
                                    pPassbookVisitorBirthdate)

        'Check that the array is large enough for a new passbook,
        'if not, grow the array.
        If _numPassbooks >= _maxPassbooks Then
            _maxPassbooks += mDEFAULT_ARRAY_GROWTH_INCREMENT
            ReDim Preserve mPassbooks(_maxPassbooks - 1)
        End If

        ' Add the Passbook to the array in the correct index
        Try
            _ithPassbook(_numPassbooks) = newPassbook
        Catch ex As Exception
            Throw New IndexOutOfRangeException
        End Try

        ' Increase the Passbook count
        _numPassbooks += 1

        ' Raise event
        RaiseEvent ThemePark_PassbookAdded(
            Me,
            New ThemePark_EventArgs_PassbookAdded(
                newPassbook
                )
            ) 'RaiseEvent

        Return newPassbook

    End Function '_addPassbook()


    ''' <summary>
    ''' Returns the index number of the passbook feature with the specific passbook feature ID
    ''' Returns -1 if not found
    ''' </summary>
    ''' <param name="pPassbookFeatureID"></param>
    ''' <returns></returns>
    Private Function _findPassbookFeature(ByVal pPassbookFeatureID As String) _
        As Integer

        'Declare variables
        Dim i As Integer = 0

        'Loop through existing Passbooks and see if we have a match

        For i = 0 To _numPassbookFeatures - 1
            If ithPassbookFeature(i).passbookFeatureID = pPassbookFeatureID Then
                Return i
            End If
        Next

        ' if nothing found, return -1
        Return -1

    End Function '_findPassbook


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

        'Declare Variables
        Dim newPassbookFeature As PassbookFeature

        ' Call the specialty constructor
        newPassbookFeature = New PassbookFeature(pPassbookFeatureID,
                                                  pQtyPurchased,
                                                  pPassbookFeatureAmt,
                                                  pPassbook,
                                                  pFeature,
                                                  pQtyRemaining)

        'Check that the array is large enough for a new passbook feature
        'if not, grow the array.
        If _numPassbookFeatures >= _maxPassbookFeatures Then
            _maxPassbookFeatures += mDEFAULT_ARRAY_GROWTH_INCREMENT
            ReDim Preserve mPassbookFeatures(_maxPassbookFeatures - 1)
        End If

        ' Add the passbook feature to the array in the correct index
        Try
            _ithPassbookFeature(_numPassbookFeatures) = newPassbookFeature
        Catch ex As Exception
            Throw New IndexOutOfRangeException
        End Try

        ' Increase the Feature count
        _numPassbookFeatures += 1

        ' Raise event
        RaiseEvent ThemePark_PassbookFeatureAdded(
            Me,
            New ThemePark_EventArgs_PassbookFeatureAdded(
                newPassbookFeature
                )
            ) 'RaiseEvent

        ' Return Ref to Object
        Return newPassbookFeature

    End Function '_addPassbookFeature()


    ''' <summary>
    ''' Returns the index number of the used passbook feature with the specific used passbook feature ID
    ''' Returns -1 if not found
    ''' </summary>
    ''' <param name="pUsedPassbookFeatureID"></param>
    ''' <returns></returns>
    Private Function _findUsedPassbookFeature(ByVal pUsedPassbookFeatureID As String) _
        As Integer

        'Declare variables
        Dim i As Integer = 0

        'Loop through existing Used Passbook Features and see if we have a match

        For i = 0 To _numUsedFeatures - 1
            If ithUsedFeature(i).usedFeatureID = pUsedPassbookFeatureID Then
                Return i
            End If
        Next

        ' if nothing found, return -1
        Return -1

    End Function '_findUsedPassbookFeature


    ''' <summary>
    ''' Creates a new UsedFeature object, increases usedFeature count
    ''' </summary>
    Private Function _addUsedFeature(ByVal pUsedFeatureID As String,
                                     ByVal pUsedPassbookFeature As PassbookFeature,
                                     ByVal pUsedDate As Date,
                                     ByVal pUsedLocation As String,
                                     ByVal pQtyUsed As Decimal) _
        As UsedFeature

        'Declare variables
        Dim newUsedFeature As UsedFeature

        ' Call the specialty constructor
        newUsedFeature = New UsedFeature(pUsedFeatureID,
                                          pUsedPassbookFeature,
                                          pUsedDate,
                                          pUsedLocation,
                                          pQtyUsed)

        'Check that the array is large enough for a new used feature,
        'if not, grow the array.
        If _numUsedFeatures >= _maxUsedFeatures Then
            _maxUsedFeatures += mDEFAULT_ARRAY_GROWTH_INCREMENT
            ReDim Preserve mUsedFeatures(_maxUsedFeatures - 1)
        End If

        ' Add the used feature to the array in the correct index
        Try
            _ithUsedFeature(_numUsedFeatures) = newUsedFeature
        Catch ex As Exception
            Throw New IndexOutOfRangeException
        End Try

        ' Increase the Used Feature count
        _numUsedFeatures += 1

        ' Need to decrease the Feature qty in matching Passbook
        pUsedPassbookFeature.passbookFeatureQtyRemaining =
            pUsedPassbookFeature.passbookFeatureQtyRemaining - pQtyUsed

        ' Raise event
        RaiseEvent ThemePark_UsedFeatureAdded(
            Me,
            New ThemePark_EventArgs_UsedFeatureAdded(
                newUsedFeature
                )
            ) 'RaiseEvent

        Return newUsedFeature

    End Function '_addUsedFeature()

    ''' <summary>
    ''' Returns a Theme Park object in String form.
    ''' </summary>
    ''' <returns>Theme Park object as String</returns>
    Private Function _toString() As String

        Dim tmpString As String
        tmpString = "( THEMEPARK: ParkName=" & mParkName _
            & ", numCustomers=" & mNumCustomers.ToString _
            & "/ maxCustomers=" & mMaxCustomers.ToString _
            & ", numPassbooks=" & mNumPassbooks.ToString _
            & "/ maxPassbooks=" & mMaxPassbooks.ToString _
            & ", numFeatures=" & mNumFeatures.ToString _
            & "/ maxFeatures=" & mMaxFeatures.ToString _
            & ", numPassbookFeatures=" & mNumPassbookFeatures.ToString _
            & "/ maxPassbookFeatures=" & mMaxPassbookFeatures.ToString _
            & ", numUsedFeatures=" & mNumUsedFeatures.ToString _
            & "/ maxUsedFeatures=" & mMaxUsedFeatures.ToString _
            & " )"

        Return tmpString

    End Function '_toString()

    ''' <summary>
    ''' Updates a new PassbookFeature object, increases passbookFeature count
    ''' </summary>
    Private Function _updatePassbookFeature(ByVal pPassbookFeature As PassbookFeature,
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

    ''' <summary>
    ''' Calculates the most popular feature based on occurances
    ''' </summary>
    ''' <returns></returns>
    Private Function _calcTopFeature() _
        As String

        Dim CurrentTopFeature As Feature
        Dim countTopFeature As Integer = 0
        Dim countTempFeature As Integer
        Dim i As Integer

        CurrentTopFeature = Nothing

        'Initialize
        countTopFeature = 0

        For i = 0 To mNumFeatures - 1
            countTempFeature = 0
            For j = 0 To mNumPassbookFeatures - 1
                'loop through existing passbook features and count occurances
                If ithPassbookFeature(j).passbookFeatureID = ithFeature(i).featureID Then
                    countTempFeature += 1
                End If
            Next

            'Now loop through the used features, and count occurence there
            For j = 0 To mNumUsedFeatures - 1
                If ithUsedFeature(j).usedPassbookFeature.feature.featureID = ithFeature(i).featureID Then
                    countTempFeature += 1
                End If
            Next

            'If the current count is more than the saved count, update the feature name
            'and count
            If countTempFeature > countTopFeature Then
                CurrentTopFeature = ithFeature(i)
                countTopFeature = countTempFeature
            End If

        Next i

        If countTopFeature = 0 Then
            Return "n.a"
        End If

        Return CurrentTopFeature.featureID

    End Function '_calcTopFeature()

    ''' <summary>
    ''' Calculates the number of birthdays today
    ''' </summary>
    ''' <returns></returns>
    Private Function _calcBirthdaysInMonth() As Integer

        Dim numBirthdays As Integer
        Dim curPassbook As Passbook
        Dim i As Integer

        'We have to have a passbook in the first place
        If mNumPassbooks = 0 Then
            Return 0
        End If

        numBirthdays = 0
        For i = 0 To mNumPassbooks - 1
            curPassbook = ithPassbook(i)
            If Month(curPassbook.passbookVisitorBirthdate) = Month(Now) Then
                numBirthdays += 1
            End If
        Next i
    End Function '_calcBirthdaysInMonth()

    ''' <summary>
    ''' Calculates Average Passbook Vistor Age
    ''' </summary>
    ''' <returns></returns>
    Private Function _calcAvgAge() As Integer

        Dim totalAge As Integer = 0
        Dim i As Integer

        'Protect against divide by zero
        If mNumPassbooks = 0 Then
            Return 0
        End If

        For i = 0 To mNumPassbooks - 1
            totalAge = totalAge + ithPassbook(i).age
        Next

        Return (CInt(totalAge / mNumPassbooks))
    End Function '_calcAvgAge()

    ''' <summary>
    ''' Calculates the average number of passbooks per customer
    ''' </summary>
    ''' <returns></returns>
    Private Function _calcAvgPBperCust() As Double
        'Protect against divide by zero
        If mNumCustomers = 0 Then
            Return 0.0
        End If

        Return (mNumPassbooks / mNumCustomers)
    End Function '_calcAvgPBperCust()

    ''' <summary>
    ''' Calculates the sum, in dollars of unused passbook features
    ''' </summary>
    ''' <returns></returns>
    Private Function _calcSumUnusedPBF() As Double

        Dim dollarTotal As Double = 0
        Dim currentPBF As PassbookFeature
        Dim i As Integer

        If mNumPassbookFeatures = 0 Then
            Return dollarTotal
        End If

        For i = 0 To mNumPassbookFeatures - 1
            currentPBF = ithPassbookFeature(i)
            If currentPBF.passbook.isChild Then
                dollarTotal += currentPBF.passbookFeatureQtyRemaining * currentPBF.feature.featureChildPrice
            Else
                dollarTotal += currentPBF.passbookFeatureQtyRemaining * currentPBF.feature.featureAdultPrice
            End If
        Next

        Return dollarTotal

    End Function '_calcSumUnusedPBF()

    ''' <summary>
    ''' Returns the average unused passbook feature balance per passbook
    ''' </summary>
    ''' <returns></returns>
    Private Function _calcAvgUnusedPBFBal() As Double

        Dim totalUnusedFeatureBal As Double = 0

        'Protect against divide by zero
        If mNumPassbooks = 0 Then
            Return 0
        End If

        Return (_calcSumUnusedPBF() / mNumPassbooks)
    End Function '_calcAvgUnusedPBFBal() 

    ''' <summary>
    ''' Returns the ratio of total dollars used / total dollars purchased
    ''' </summary>
    ''' <returns></returns>
    Private Function _calcPercentFeatUse() As Double

        Dim totalDollarsPurchased As Double = 0
        Dim totalDollarsUsed As Double = 0
        Dim thePassbookFeature As PassbookFeature
        Dim theUsedFeature As UsedFeature
        Dim i As Integer

        'Calculate total dollars purchased
        For i = 0 To mNumPassbookFeatures - 1
            thePassbookFeature = ithPassbookFeature(i)
            totalDollarsPurchased += thePassbookFeature.passbookFeatureAmt
        Next i

        'Calculate the total dollars used
        For i = 0 To mNumUsedFeatures - 1
            theUsedFeature = ithUsedFeature(i)
            If theUsedFeature.usedPassbookFeature.passbook.isChild Then
                totalDollarsUsed += theUsedFeature.usedQty * theUsedFeature.usedPassbookFeature.feature.featureChildPrice
            Else
                totalDollarsUsed += theUsedFeature.usedQty * theUsedFeature.usedPassbookFeature.feature.featureAdultPrice
            End If
        Next

        'Check for divide by zero
        If totalDollarsPurchased = 0 Then
            Return 0
        End If

        Return 100 * (totalDollarsUsed / totalDollarsPurchased)
    End Function '_calcPercentFeatureUse()

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


