Attribute VB_Name = "Module1"
Option Explicit
Sub main_process()
Call init
Call create_functions
Call run_impact_analysis

End Sub

Sub init()
' define variables, naming array formula
    ProductMix.Activate
    Range("B4").Name = "labor_unit_cost"
    Range("B5").Name = "metal_unit_cost"
    Range("b6").Name = "glass_unit_cost"
    
    Range("B9:E9").Name = "labor_per_frame"
    Range("B10:E10").Name = "metal_per_frame"
    Range("B11:E11").Name = "glass_per_frame"
    Range("B12:E12").Name = "unit_selling_price"
    
    Range("B16:E16").Name = "produced"
    Range("B18:E18").Name = "max_sales"
    
    Range("B21").Name = "labor_used"
    Range("B22").Name = "metal_used"
    Range("B23").Name = "glass_used"
    Range("D21").Name = "res_avail_labor"
    Range("D22").Name = "res_avail_metal"
    Range("D23").Name = "res_avail_glass"
    
    
    Range("B26:E26").Name = "revenue"
    Range("B28:E28").Name = "labor_cost"
    Range("B29:E29").Name = "glass_cost"
    Range("B30:E30").Name = "metal_cost"
    Range("B33:E33").Name = "total_cost"
    Range("E34").Name = "max_profit"
    
    Range("J32:J39").Name = "profit_values"
    Range("K32:N39").Name = "impact_analysis"
End Sub
Sub create_functions()
    'resource constraints; single cell calculation
    Range("labor_used").Formula = "=sumproduct(produced,labor_per_frame)"
    Range("metal_used").Formula = "=sumproduct(produced,metal_per_frame)"
    Range("glass_used").Formula = "=sumproduct(produced,glass_per_frame)"
    
    'revenue - array
    Range("revenue").FormulaArray = "=Produced * unit_selling_price"
    
    'cost - array
    Range("labor_cost").FormulaArray = "=produced * labor_unit_cost * labor_per_frame"
    Range("glass_cost").FormulaArray = "=produced * glass_unit_cost * glass_per_frame"
    Range("metal_cost").FormulaArray = "=produced * metal_unit_cost * metal_per_frame"
    
    Range("total_cost").FormulaArray = "=labor_cost + glass_cost + metal_cost"

    'max profit
    Range("max_profit").FormulaArray = "=sum(revenue-total_cost)"
End Sub

Sub run_impact_analysis()
    Range("impact_analysis").Clear
    Dim cell As Range
   ' Range("impact_analysis").Clear
    For Each cell In Range("profit_values")
        Call run_simulation(cell.value)
        Range("produced").Copy cell.Offset(0, 1)
        
    Next cell
    
End Sub

Sub run_simulation(value As Integer)
    'clear
    solverreset
        SolverOk SetCell:=Range("Max_Profit"), MaxMinVal:=3, valueof:=value, ByChange:=Range("produced")

        SolverAdd CellRef:=Range("produced"), Relation:=1, FormulaText:=Range("max_sales")
        SolverAdd CellRef:=Range("labor_used"), Relation:=1, FormulaText:=Range("res_avail_labor")
        SolverAdd CellRef:=Range("metal_used"), Relation:=1, FormulaText:=Range("res_avail_metal")
        SolverAdd CellRef:=Range("glass_used"), Relation:=1, FormulaText:=Range("res_avail_glass")
        
        'dont have to confirm each time
        SolverSolve True
    
End Sub

