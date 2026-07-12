"""Direct spending, debt, forecast, merchant and financial-DNA analysis."""
from __future__ import annotations
import re
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path
from statistics import median, pstdev
from typing import Any
import yaml
from src.extract import HistoricalExtractionResult
from src.kpi import AnnualKPI, CurrentSnapshot
from src.transform import CategoryRegistry
ZERO=Decimal('0'); TWELVE=Decimal('12'); FIFTY_TWO=Decimal('52')
@dataclass(frozen=True,slots=True)
class MerchantSummary:
    original_label:str; mapped_category_id:str|None; mapped_category:str|None; suggested_category_id:str|None; suggested_category:str|None; rule_id:str|None; confidence:str; opportunity_group:str|None; transaction_count:int; total_amount:Decimal; latest_year:int; review_status:str
@dataclass(frozen=True,slots=True)
class SpendingOpportunity:
    rank:int; category_id:str; opportunity:str; essentiality:str; annual_spend:Decimal; monthly_spend:Decimal; save_25:Decimal; save_50:Decimal; save_100:Decimal; loc_weeks_saved_50:int; loc_weeks_saved_100:int; recommendation:str; direct_insight:str
@dataclass(frozen=True,slots=True)
class DebtScenario:
    scenario:str; weekly_payment:Decimal; annual_redirect:Decimal; payoff_weeks:int; payoff_months:Decimal; interest_cost:Decimal; interest_saved:Decimal; weeks_saved:int; ending_balance_12m:Decimal
@dataclass(frozen=True,slots=True)
class ForecastMonth:
    scenario:str; month_number:int; income:Decimal; operating_expenses:Decimal; discretionary_reduction:Decimal; wealth_building:Decimal; loc_payment:Decimal; loc_interest:Decimal; unallocated_margin:Decimal; ending_loc:Decimal; ending_savings:Decimal
@dataclass(frozen=True,slots=True)
class ForecastSummary:
    scenario:str; reduction_rate:Decimal; annual_discretionary_savings:Decimal; loc_payoff_month:int|None; ending_loc:Decimal; ending_savings:Decimal; annual_unallocated_margin:Decimal
@dataclass(frozen=True,slots=True)
class CategoryDNA:
    category:str; years_observed:int; eligible_years:int; persistence_ratio:Decimal; first_year:int; latest_year:int; average_spend:Decimal; median_spend:Decimal; latest_spend:Decimal; latest_share_of_income:Decimal|None; annual_trend:Decimal; volatility:Decimal|None; peak_year:int; peak_spend:Decimal; pattern:str
@dataclass(frozen=True,slots=True)
class DNAFinding:
    rank:int; theme:str; finding:str; evidence:str; implication:str
@dataclass(frozen=True,slots=True)
class CategoryHistoryRow:
    year:int; coverage:str; comparison_eligible:bool; level:str; category_id:str; category:str; amount:Decimal; share_of_income:Decimal|None; change_vs_prior:Decimal|None; change_vs_prior_ratio:Decimal|None
@dataclass(frozen=True,slots=True)
class DecisionReport:
    latest_year:int; assumptions:dict[str,Decimal|int]; merchants:tuple[MerchantSummary,...]; opportunities:tuple[SpendingOpportunity,...]; debt_scenarios:tuple[DebtScenario,...]; forecast_months:tuple[ForecastMonth,...]; forecast_summaries:tuple[ForecastSummary,...]; category_dna:tuple[CategoryDNA,...]; dna_findings:tuple[DNAFinding,...]; category_history:tuple[CategoryHistoryRow,...]
class DecisionIntelligenceEngine:
    def __init__(self, registry:CategoryRegistry, config_file:str|Path)->None:
        self.registry=registry; self.entries={str(e['category_id']):e for e in registry.entries()}
        with Path(config_file).open('r',encoding='utf-8') as f: config:Any=yaml.safe_load(f)
        if not isinstance(config,dict): raise ValueError('Decision intelligence configuration must be a mapping.')
        self.config=config; self.assumptions=dict(config.get('assumptions',{})); self.opportunity_config=tuple(config.get('opportunities',[])); self.merchant_rules=tuple(config.get('merchant_rules',[]))
    def analyze(self,extraction:HistoricalExtractionResult,annual_kpis:tuple[AnnualKPI,...],snapshot:CurrentSnapshot)->DecisionReport:
        complete=sorted((k for k in annual_kpis if k.coverage_status=='Complete'),key=lambda k:k.year)
        if not complete: raise ValueError('Decision intelligence requires one complete annual year.')
        latest=complete[-1]; merchants=self._merchant_summaries(extraction); opps=self._opportunities(extraction,latest,snapshot); debt=self._debt_scenarios(snapshot,opps); fm,fs=self._forecast(extraction,latest,snapshot,opps); hist=self._category_history(extraction,annual_kpis); dna=self._category_dna(hist,annual_kpis); findings=self._dna_findings(dna,annual_kpis,latest)
        return DecisionReport(latest.year,{'loc_annual_interest_rate':Decimal(str(self.assumptions.get('loc_annual_interest_rate',0.065))),'loc_weekly_payment':Decimal(str(self.assumptions.get('loc_weekly_payment',500))),'forecast_months':int(self.assumptions.get('forecast_months',12)),'emergency_target_months':Decimal(str(self.assumptions.get('emergency_target_months',3)))},tuple(merchants),tuple(opps),tuple(debt),tuple(fm),tuple(fs),tuple(dna),tuple(findings),tuple(hist))
    @staticmethod
    def _norm(v:str)->str: return ' '.join(re.sub(r'[^a-z0-9&]+',' ',v.casefold()).split())
    def _match_rule(self,label:str)->dict[str,Any]|None:
        n=self._norm(label)
        for r in self.merchant_rules:
            if any(self._norm(str(t)) in n for t in r.get('contains',[])): return dict(r)
        return None
    def _merchant_summaries(self,extraction)->list[MerchantSummary]:
        mapped={}
        for s in extraction.sheets:
            for rec in s.result.records:
                tx=rec.transaction; mapped[(tx.source_sheet,tx.source_cell)]=tx.category_id
        grouped={}
        for s in extraction.sheets:
            for src in s.result.source_rows:
                mid=mapped.get((src.source_sheet,src.source_cell)); rule=self._match_rule(src.original_name); sid=str(rule['suggested_category_id']) if rule else None; rid=str(rule['rule_id']) if rule else None; key=(src.original_name,mid,rid)
                g=grouped.setdefault(key,{'count':0,'amount':ZERO,'latest':0,'rule':rule,'sid':sid}); g['count']+=1; g['amount']+=src.amount; g['latest']=max(g['latest'],int(src.source_sheet[:4]))
        out=[]
        for (label,mid,rid),g in grouped.items():
            rule=g['rule']; sid=g['sid']; mname=str(self.entries[mid]['display_name']) if mid else None; sname=str(self.entries[sid]['display_name']) if sid in self.entries else None; status='Mapped' if mid else ('Suggested — review before mapping' if sid else 'Unresolved')
            out.append(MerchantSummary(label,mid,mname,sid,sname,rid,str(rule.get('confidence','None')) if rule else 'None',str(rule.get('opportunity_group')) if rule and rule.get('opportunity_group') else None,int(g['count']),g['amount'],int(g['latest']),status))
        return sorted(out,key=lambda x:(x.review_status=='Mapped',-x.total_amount,x.original_label.casefold()))
    def _latest_totals(self,extraction,year:int)->dict[str,Decimal]:
        sh=next(s for s in extraction.sheets if s.year==year); totals={}
        for tx in list(sh.result.transfers)+list(sh.result.variable_expenses)+list(sh.result.fixed_expenses): totals[tx.category_id]=totals.get(tx.category_id,ZERO)+tx.amount
        return totals
    @staticmethod
    def _simulate(balance:Decimal,rate:Decimal,payment:Decimal,max_weeks:int=1040)->tuple[int,Decimal,Decimal]:
        if balance<=0:return 0,ZERO,ZERO
        wr=rate/FIFTY_TWO; cur=balance; interest=ZERO; end12=balance
        for w in range(1,max_weeks+1):
            i=cur*wr; interest+=i; cur+=i; cur=max(ZERO,cur-min(payment,cur))
            if w==52:end12=cur
            if cur<=Decimal('0.005'):
                if w<52:end12=ZERO
                return w,interest,end12
        return max_weeks,interest,end12
    def _opportunities(self,extraction,latest:AnnualKPI,snapshot:CurrentSnapshot)->list[SpendingOpportunity]:
        totals=self._latest_totals(extraction,latest.year); rate=Decimal(str(self.assumptions.get('loc_annual_interest_rate',0.065))); base=Decimal(str(self.assumptions.get('loc_weekly_payment',500))); basew,_,_=self._simulate(snapshot.line_of_credit_balance,rate,base); temp=[]
        for c in self.opportunity_config:
            cid=str(c['category_id']); annual=totals.get(cid,ZERO)
            if annual<=0: continue
            hw,_,_=self._simulate(snapshot.line_of_credit_balance,rate,base+annual*Decimal('0.5')/FIFTY_TWO); fw,_,_=self._simulate(snapshot.line_of_credit_balance,rate,base+annual/FIFTY_TWO); name=str(c['name']); direct=f"You spent ${annual:,.0f} on {name.lower()} in {latest.year}. A 50% reduction would free ${annual*Decimal('0.5'):,.0f}/yr; stopping entirely would free ${annual:,.0f}/yr."
            temp.append((annual,c,hw,fw,direct))
        temp.sort(key=lambda x:(-x[0],int(x[1].get('priority',99))))
        out=[]
        for rank,(annual,c,hw,fw,direct) in enumerate(temp,1): out.append(SpendingOpportunity(rank,str(c['category_id']),str(c['name']),str(c.get('essentiality','Discretionary')),annual,annual/TWELVE,annual*Decimal('0.25'),annual*Decimal('0.5'),annual,max(0,basew-hw),max(0,basew-fw),str(c.get('recommendation','Review spending.')),direct))
        return out
    def _debt_scenarios(self,snapshot,opps)->list[DebtScenario]:
        rate=Decimal(str(self.assumptions.get('loc_annual_interest_rate',0.065))); base=Decimal(str(self.assumptions.get('loc_weekly_payment',500))); bw,bi,b12=self._simulate(snapshot.line_of_credit_balance,rate,base); total=sum((o.annual_spend for o in opps),ZERO); defs=[('Current $500/week plan',ZERO),('Redirect 25% of opportunities',total*Decimal('0.25')),('Redirect 50% of opportunities',total*Decimal('0.5')),('Redirect 100% of opportunities',total)]+[(f'Eliminate {o.opportunity}',o.annual_spend) for o in opps[:3]]; out=[]
        for name,redir in defs:
            pay=base+redir/FIFTY_TWO; w,i,e12=self._simulate(snapshot.line_of_credit_balance,rate,pay); out.append(DebtScenario(name,pay,redir,w,Decimal(w)*TWELVE/FIFTY_TWO,i,max(ZERO,bi-i),max(0,bw-w),e12))
        return out
    def _forecast(self,extraction,latest,snapshot,opps):
        months=int(self.assumptions.get('forecast_months',12)); annual_opp=sum((o.annual_spend for o in opps),ZERO); rate=Decimal(str(self.assumptions.get('loc_annual_interest_rate',0.065)))/TWELVE; basepay=Decimal(str(self.assumptions.get('loc_weekly_payment',500)))*FIFTY_TWO/TWELVE; sh=next(s for s in extraction.sheets if s.year==latest.year); sav=sum((t.amount for t in sh.result.transfers if t.category_id=='TRF001'),ZERO)/TWELVE; income=latest.true_income/TWELVE; operating=latest.known_operating_expenses/TWELVE; wealth=latest.wealth_building/TWELVE; rows=[]; sums=[]
        for sc in self.assumptions.get('forecast_scenarios',[]):
            name=str(sc['name']); red=Decimal(str(sc['reduction_rate'])); mred=annual_opp*red/TWELVE; loc=snapshot.line_of_credit_balance; savings=snapshot.savings_cash; payoff=None; mt=ZERO
            for m in range(1,months+1):
                interest=loc*rate if loc>0 else ZERO; loc+=interest; desired=basepay+mred; actual=min(loc,desired); loc-=actual; excess=max(ZERO,desired-actual)
                if loc<=Decimal('0.005'): loc=ZERO; payoff=payoff or m
                savings+=sav+excess; adj=max(ZERO,operating-mred); margin=income-adj-wealth-actual; mt+=margin; rows.append(ForecastMonth(name,m,income,adj,mred,wealth,actual,interest,margin,loc,savings))
            sums.append(ForecastSummary(name,red,annual_opp*red,payoff,loc,savings,mt))
        return rows,sums
    def _category_history(self,extraction,annual_kpis):
        kp={k.year:k for k in annual_kpis}; prior={}; out=[]
        for sh in sorted(extraction.sheets,key=lambda s:s.year):
            k=kp[sh.year]; master={}; cats={}
            for tx in list(sh.result.transfers)+list(sh.result.variable_expenses)+list(sh.result.fixed_expenses):
                e=self.entries[tx.category_id]
                if e['category_type']=='Transfer': continue
                cats[tx.category_id]=cats.get(tx.category_id,ZERO)+tx.amount; m=str(e['master_category']); master[m]=master.get(m,ZERO)+tx.amount
            candidates=[('Master',n,n,a) for n,a in master.items()]+[('Category',cid,str(self.entries[cid]['display_name']),a) for cid,a in cats.items()]
            for level,cid,name,amt in sorted(candidates):
                key=(level,cid); prev=prior.get(key); ch=amt-prev if prev is not None else None; ratio=ch/prev if ch is not None and prev not in (None,ZERO) else None; out.append(CategoryHistoryRow(sh.year,k.coverage_status,k.comparison_eligible,level,cid,name,amt,amt/k.true_income if k.true_income else None,ch,ratio)); prior[key]=amt
        return out
    @staticmethod
    def _slope(values:list[tuple[int,Decimal]])->Decimal:
        if len(values)<2:return ZERO
        xs=[Decimal(y) for y,_ in values]; ys=[a for _,a in values]; xm=sum(xs,ZERO)/Decimal(len(xs)); ym=sum(ys,ZERO)/Decimal(len(ys)); den=sum(((x-xm)**2 for x in xs),ZERO); return ZERO if den==0 else sum(((x-xm)*(y-ym) for x,y in zip(xs,ys)),ZERO)/den
    def _category_dna(self,history,annual_kpis):
        yrs=sorted(k.year for k in annual_kpis if k.comparison_eligible); by={}
        for r in history:
            if r.level=='Master' and r.comparison_eligible: by.setdefault(r.category,[]).append(r)
        out=[]
        for cat,rows in by.items():
            vals={r.year:r.amount for r in rows}; series=[(y,vals.get(y,ZERO)) for y in yrs]; nz=[(y,a) for y,a in series if a>0]
            if not nz:continue
            amounts=[a for _,a in series]; nza=[a for _,a in nz]; avg=sum(amounts,ZERO)/Decimal(len(amounts)); med=Decimal(str(median([float(a) for a in amounts]))); nza_avg=sum(nza,ZERO)/Decimal(len(nza)); std=Decimal(str(pstdev([float(a) for a in nza]))) if len(nza)>1 else ZERO; vol=std/nza_avg if nza_avg else None; slope=self._slope(series); pers=Decimal(len(nz))/Decimal(len(yrs)); peak_y,peak=max(series,key=lambda x:x[1]); rel=slope/avg if avg else ZERO
            if pers<Decimal('0.50'):pattern='Intermittent'
            elif rel>=Decimal('0.05'):pattern='Rising'
            elif rel<=Decimal('-0.05'):pattern='Declining'
            elif vol is not None and vol>=Decimal('0.60'):pattern='Volatile'
            elif pers>=Decimal('0.80'):pattern='Persistent'
            else:pattern='Intermittent'
            lr=max(rows,key=lambda r:r.year); out.append(CategoryDNA(cat,len(nz),len(yrs),pers,nz[0][0],lr.year,avg,med,lr.amount,lr.share_of_income,slope,vol,peak_y,peak,pattern))
        return sorted(out,key=lambda x:x.latest_spend,reverse=True)
    def _dna_findings(self,profiles,annual_kpis,latest):
        out=[]; rising=sorted((p for p in profiles if p.pattern=='Rising'),key=lambda p:p.annual_trend,reverse=True); persistent=sorted((p for p in profiles if p.persistence_ratio>=Decimal('0.80')),key=lambda p:p.latest_spend,reverse=True); volatile=sorted((p for p in profiles if p.pattern=='Volatile' and p.average_spend>=Decimal('2000') and p.persistence_ratio>=Decimal('0.60')),key=lambda p:p.volatility or ZERO,reverse=True)
        if rising:
            p=rising[0]; out.append(DNAFinding(1,'Spending drift',f'{p.category} has the strongest sustained upward trend.',f'The fitted trend is approximately ${p.annual_trend:,.0f} per year across comparison-eligible history.','This category deserves a recurring budget rule rather than a one-time cut.'))
        if persistent:
            p=persistent[0]; out.append(DNAFinding(2,'Structural cost',f'{p.category} is the largest persistent spending category.',f'It appeared in {p.years_observed} of {p.eligible_years} comparable years and was ${p.latest_spend:,.0f} in {p.latest_year}.','Long-term improvement depends more on contract, housing or service decisions than weekly restraint.'))
        if volatile:
            p=volatile[0]; out.append(DNAFinding(3,'Irregular pressure',f'{p.category} is the most volatile major category.',f'Its volatility ratio is {(p.volatility or ZERO)*Decimal("100"):.0f}% and its peak was ${p.peak_spend:,.0f} in {p.peak_year}.','A sinking fund is more appropriate than treating this spending as a surprise.'))
        comp=sorted((k for k in annual_kpis if k.comparison_eligible),key=lambda k:k.year)
        if comp:
            avg=sum((k.wealth_building_rate or ZERO for k in comp),ZERO)/Decimal(len(comp)); out.append(DNAFinding(4,'Wealth habit','Wealth building is a recurring household behaviour, not a one-year event.',f'The comparable-year average was {avg*Decimal("100"):.1f}% of income; {latest.year} was {(latest.wealth_building_rate or ZERO)*Decimal("100"):.1f}%.','Protecting the habit while redirecting discretionary spending is preferable to stopping long-term saving.'))
        return out
