from decimal import Decimal
from pathlib import Path
from src.decision import DecisionIntelligenceEngine
from src.extract import CurrentExtractionResult, ExtractedRecord, HistoricalExtractionResult, SheetExtraction, SourceRow, UnknownCategory
from src.kpi import AnnualKPI, CurrentSnapshot
from src.models import Transaction
from src.transform import CategoryRegistry
ROOT=Path(__file__).resolve().parents[1]
def annual(year,income,fixed,variable,wealth):
    i=Decimal(income); f=Decimal(fixed); v=Decimal(variable); w=Decimal(wealth); known=f+v
    return AnnualKPI(year,'Complete',26,i,Decimal('0'),f,v,known,Decimal('72000'),Decimal('9000'),w,Decimal('0'),Decimal('0'),Decimal('100'),f/i,known/i,w/i,w/i,(i-f)/i,Decimal('.2'),Decimal('.08'),Decimal('.05'),Decimal('.04'),Decimal('.98'),True)
def snap():
    return CurrentSnapshot('A & L',2025,Decimal('1000000'),Decimal('500000'),Decimal('500000'),Decimal('120000'),Decimal('6000'),Decimal('450000'),Decimal('450000'),Decimal('17500'),Decimal('54000'),Decimal('.9'),Decimal('.08'),Decimal('38'),'Moderate')
def mk_sheet(year,alcohol,dining,housing):
    vals=[('FOD003','Alcohol',alcohol,'A2','variable_expenses'),('FOD002','Eating out',dining,'C2','variable_expenses'),('HOU001','Mortgage',housing,'E2','fixed_expenses'),('TRF001','Savings','12000','C3','transfers')]
    recs=[]; src=[]
    for cid,label,amt,cell,section in vals:
        tx=Transaction(year,'Jan 1 - Jan 14',cid,label,Decimal(amt),str(year),cell)
        recs.append(ExtractedRecord(section,tx)); src.append(SourceRow(tx.period,label,tx.amount,str(year),cell))
    unk=UnknownCategory('Jan 1 - Jan 14','Hockey tickets',Decimal('250'),str(year),'A7')
    src.append(SourceRow(unk.period,unk.original_name,unk.amount,unk.source_sheet,unk.source_cell))
    return SheetExtraction(str(year),'current',CurrentExtractionResult(tuple(recs),(unk,),tuple(src),('Jan 1 - Jan 14',)))
def build_report():
    reg=CategoryRegistry(ROOT/'config/categories.yaml'); eng=DecisionIntelligenceEngine(reg,ROOT/'config/decision_intelligence.yaml')
    ext=HistoricalExtractionResult((mk_sheet(2023,'1800','2200','28000'),mk_sheet(2024,'2400','3000','30000'),mk_sheet(2025,'2720','3088','33000')))
    anns=(annual(2023,'200000','42000','30000','20000'),annual(2024,'220000','50000','32000','24000'),annual(2025,'215000','60000','28000','23000'))
    return eng.analyze(ext,anns,snap())
def test_direct_opportunity_and_merchant_suggestion():
    r=build_report(); a=next(x for x in r.opportunities if x.category_id=='FOD003'); assert a.annual_spend==Decimal('2720'); assert 'You spent $2,720 on alcohol' in a.direct_insight
    m=next(x for x in r.merchants if x.original_label=='Hockey tickets'); assert m.mapped_category_id is None; assert m.suggested_category_id=='LIF001'
def test_debt_forecast_and_dna():
    r=build_report(); base=r.debt_scenarios[0]; half=next(x for x in r.debt_scenarios if '50%' in x.scenario); assert half.payoff_weeks<base.payoff_weeks; assert len(r.forecast_months)==36; assert len(r.forecast_summaries)==3; assert r.dna_findings; assert any(x.category_id=='FOD003' for x in r.category_history)
