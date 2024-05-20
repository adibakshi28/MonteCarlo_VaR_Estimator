Function CalculateAnalyticalVaR(pricesRange As Range, volatilitiesRange As Range, holdingsRange As Range, correlationMatrixRange As Range, significanceLevel As Double) As Double
    Dim alpha_percentile As Double
    Dim sigmaV As Double
    Dim portfolioVariance As Double
    Dim i As Integer, j As Integer
    Dim prices As Variant, volatilities As Variant, holdings As Variant, correlationMatrix As Variant
    
    ' Convert ranges to arrays for faster processing
    prices = pricesRange.Value
    volatilities = volatilitiesRange.Value
    holdings = holdingsRange.Value
    correlationMatrix = correlationMatrixRange.Value
    
    ' Calculate alpha for the given percentile of the normal distribution
    alpha_percentile = Application.WorksheetFunction.NormSInv(significanceLevel)
    
    ' Calculate the portfolio variance
    portfolioVariance = 0
    For i = 1 To UBound(prices, 1)
        For j = 1 To UBound(prices, 1)
            portfolioVariance = portfolioVariance + _
                holdings(i, 1) * holdings(j, 1) * prices(i, 1) * prices(j, 1) * _
                volatilities(i, 1) * volatilities(j, 1) * correlationMatrix(i, j) / 252
        Next j
    Next i
    
    ' Calculate the standard deviation of the portfolio (sigmaV)
    sigmaV = Sqr(portfolioVariance)
    
    ' Calculate VaR
    CalculateAnalyticalVaR = alpha_percentile * sigmaV
End Function


' Function to perform quicksort on the array
Sub QuickSort(ByRef arr() As Double, ByVal first As Long, ByVal last As Long)
    Dim pivot As Double, temp As Double
    Dim i As Long, j As Long

    If first >= last Then Exit Sub
    
    i = first
    j = last
    pivot = arr((first + last) \ 2)
    
    While i <= j
        While arr(i) < pivot And i < last
            i = i + 1
        Wend
        While arr(j) > pivot And j > first
            j = j - 1
        Wend
        
        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Wend
    
    If first < j Then QuickSort arr, first, j
    If i < last Then QuickSort arr, i, last
End Sub



Function CholeskyDecomposition(corrMatrixRange As Range) As Variant
    Dim n As Integer
    n = corrMatrixRange.Rows.Count
    Dim R() As Double
    ReDim R(1 To n, 1 To n)

    Dim i As Integer, j As Integer, k As Integer
    Dim sum As Double

    For i = 1 To n
        For j = 1 To i
            sum = corrMatrixRange.Cells(i, j).Value
            For k = 1 To j - 1
                sum = sum - R(i, k) * R(j, k)
            Next k
            If i = j Then
                R(i, j) = Sqr(sum)
            Else
                R(i, j) = sum / R(j, j)
            End If
        Next j
    Next i

    CholeskyDecomposition = R
End Function





Function MonteCarloVaRMSE(prices As Range, volatilities As Range, holdings As Range, _
                           corrMatrixRange As Range, significanceLevel As Double, _
                           analyticalVaR As Double, numScenarios As Long, numRepeats As Integer) As Double
    
    ' Load inputs into arrays for easier manipulation
    Dim priceArray As Variant, volArray As Variant, holdingArray As Variant
    priceArray = prices.Value2
    volArray = volatilities.Value2
    holdingArray = holdings.Value2
    
    ' Convert Range to array for correlation matrix
    Dim corrMatrix() As Double
    ReDim corrMatrix(1 To corrMatrixRange.Rows.Count, 1 To corrMatrixRange.Columns.Count)
    Dim i As Long, j As Long
    For i = 1 To corrMatrixRange.Rows.Count
        For j = 1 To corrMatrixRange.Columns.Count
            corrMatrix(i, j) = corrMatrixRange.Cells(i, j).Value
        Next j
    Next i
    
    ' Cholesky decomposition to get the lower triangular matrix
    Dim L() As Double
    L = CholeskyDecomposition(corrMatrixRange)
    
    ' Set up variables for Monte Carlo simulation
    Dim simResults() As Double, simVaRs() As Double
    Dim sumSquareError As Double
    Dim zScores() As Double
    ReDim zScores(1 To UBound(priceArray, 1))
    ReDim simVaRs(1 To numRepeats)
    
    Dim k As Integer, sim As Integer
    Dim percentile As Double
    Dim deltaV As Double
    
    ' Run the Monte Carlo simulation numRepeats times
    For sim = 1 To numRepeats
        ReDim simResults(1 To numScenarios)
        
        ' Each simulation
        For i = 1 To numScenarios
            ' Generate correlated random variables
            For j = 1 To UBound(priceArray, 1)
                zScores(j) = Application.WorksheetFunction.Norm_S_Inv(Rnd())
            Next j
            
            ' Multiply by the Cholesky decomposition to get correlated z-scores
            For j = 1 To UBound(priceArray, 1)
                deltaV = 0
                For k = 1 To j
                    deltaV = deltaV + L(j, k) * zScores(k)
                Next k
                
                ' Apply the change in value
                simResults(i) = simResults(i) + holdingArray(j, 1) * priceArray(j, 1) * deltaV * volArray(j, 1) / Sqr(252)
            Next j
        Next i
        
        ' Sort the simulation results using quicksort
        QuickSort simResults, 1, numScenarios
        
        ' Determine the value at the specified percentile
        Dim percentileIndex As Long
        percentileIndex = CLng((numScenarios - 1) * significanceLevel + 1)
        simVaRs(sim) = simResults(percentileIndex)
    Next sim
    
    ' Calculate MSE
    For i = 1 To numRepeats
        sumSquareError = sumSquareError + (simVaRs(i) - analyticalVaR) ^ 2
    Next i
    
    MonteCarloVaRMSE = Sqr(sumSquareError / numRepeats)
End Function
