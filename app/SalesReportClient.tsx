"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

type Channel = string;
type Manager = string;
type StoreType = string;
type PeriodType = "current" | "prevMonth" | "prevYear";
type SalesView = "거래처별" | "브랜드별" | "담당자별" | "채널별";
type MonthStartTab = "거래처/휴면 관리" | "Target/EST 관리" | "업로드 관리";
type DrillPeriod = "prevYear" | "prevMonth" | "current" | "currentFullMonth";
type SalesStatusSortKey = "label" | "prevYearSales" | "prevYearRate" | "prevMonthSales" | "prevMonthRate" | "currentSales" | "fullMonthSales" | "timeGone" | "timeGoneGap" | "est" | "estRate" | "profitAmount" | "profitRate";
type SortDirection = "asc" | "desc";
type InactiveOrderTab = "거래처별" | "품목별";


type Store = {
  code: string;
  name: string;
  channel: Channel;
  manager: Manager;
  storeType: StoreType;
  brand: string;
  status: "거래중" | "거래종료";
};

type StoreCodeMapping = {
  id: string;
  oldCode: string;
  oldName: string;
  currentCode: string;
  currentName: string;
  memo?: string;
};

type SalesRecord = {
  id: string;
  period: PeriodType;
  refMonth: string;
  saleDate: string;
  storeCode: string;
  storeName: string;
  channel: Channel;
  manager: Manager;
  storeType: StoreType;
  brand: string;
  itemCode: string;
  itemName: string;
  quantity: number;
  salesAmount: number;
  costAmount: number;
  profitAmount: number;
  profitRate: number;
};

type TargetRecord = {
  month: string;
  amount: number;
  storeType?: StoreType;
  storeCode?: string;
  storeName?: string;
};

type EstRecord = {
  storeCode: string;
  storeName: string;
  month: string;
  amount: number;
};

type TimeConfig = {
  month: string;
  holidays: string[];
};

const CHANNELS: Channel[] = ["도매", "체인", "체인물류", "식자재마트", "제조", "권역배송", "온라인", "매장", "비매장", "기업", "매입", "본사"];
const MANAGERS: Manager[] = ["SY", "KT", "SW", "NH", "Bomi", "BM", "bomi"];
const SALES_VIEWS: SalesView[] = ["거래처별", "브랜드별", "담당자별", "채널별"];
const MONTH_TABS: MonthStartTab[] = ["거래처/휴면 관리", "Target/EST 관리", "업로드 관리"];

const initialStores: Store[] = [
  {
    "code": "385-81-04167",
    "name": "주식회사 메이트레이더스",
    "channel": "도매",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "이마트 (수출)",
    "status": "거래중"
  },
  {
    "code": "1298688941-1",
    "name": "(주)에스피씨 지에프에스_왓더버거",
    "channel": "체인",
    "manager": "KT",
    "storeType": "비매장",
    "brand": "SPC(왓더버거)",
    "status": "거래중"
  },
  {
    "code": "6018700459",
    "name": "명랑시대외식청년창업협동조합",
    "channel": "체인",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "명랑시대",
    "status": "거래중"
  },
  {
    "code": "4678601074-3",
    "name": "세이웰_왁버거(제때)",
    "channel": "체인물류",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "세이웰",
    "status": "거래중"
  },
  {
    "code": "417-81-24010",
    "name": "주식회사 해창수산",
    "channel": "도매",
    "manager": "KT",
    "storeType": "비매장",
    "brand": "주식회사 해창수산",
    "status": "거래중"
  },
  {
    "code": "237-88-02985",
    "name": "주식회사 하라에프에스",
    "channel": "도매",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "주식회사 하라에프에스",
    "status": "거래중"
  },
  {
    "code": "1368702986",
    "name": "주식회사 쓰담",
    "channel": "제조",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "주식회사 쓰담",
    "status": "거래중"
  },
  {
    "code": "119-86-48020",
    "name": "더한솔씨앤에스(주)",
    "channel": "체인",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "더한솔씨앤에스(주)",
    "status": "거래중"
  },
  {
    "code": "999-1",
    "name": "[온라인] 스마트스토어 고객",
    "channel": "온라인",
    "manager": "Bomi",
    "storeType": "비매장",
    "brand": "온라인",
    "status": "거래중"
  },
  {
    "code": "447-81-01963",
    "name": "주식회사 메이코더스",
    "channel": "도매",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "이마트 (수출)",
    "status": "거래중"
  },
  {
    "code": "4678601074-2",
    "name": "세이웰_번패티번(CJFW)",
    "channel": "체인물류",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "세이웰",
    "status": "거래중"
  },
  {
    "code": "2598501011",
    "name": "(주)푸드엔 물류센터",
    "channel": "식자재마트",
    "manager": "KT",
    "storeType": "비매장",
    "brand": "(주)푸드엔 물류센터",
    "status": "거래중"
  },
  {
    "code": "5038701038",
    "name": "주식회사 조인앤조인",
    "channel": "제조",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "주식회사 조인앤조인",
    "status": "거래중"
  },
  {
    "code": "268-88-02001",
    "name": "세븐패티버거 을지로본점",
    "channel": "체인",
    "manager": "SW",
    "storeType": "매장",
    "brand": "세븐패티버거",
    "status": "거래중"
  },
  {
    "code": "8078802372",
    "name": "주식회사 링커",
    "channel": "도매",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "주식회사 링커",
    "status": "거래중"
  },
  {
    "code": "4638501852",
    "name": "온더보더여의도 IFC점 (SRW)",
    "channel": "체인",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "온더보더",
    "status": "거래중"
  },
  {
    "code": "2028535854",
    "name": "온더보더 광화문 D타워점 (SRW)",
    "channel": "체인",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "온더보더",
    "status": "거래중"
  },
  {
    "code": "6765800749",
    "name": "더블 식스 버거",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "더블 식스 버거",
    "status": "거래중"
  },
  {
    "code": "8262002067",
    "name": "스니커스(snickers)홍대",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "스니커스(snickers)홍대",
    "status": "거래중"
  },
  {
    "code": "268-88-02002",
    "name": "세븐패티버거 압구정로데오점",
    "channel": "체인",
    "manager": "SW",
    "storeType": "매장",
    "brand": "세븐패티버거",
    "status": "거래중"
  },
  {
    "code": "4678601074",
    "name": "주식회사 세이웰",
    "channel": "체인물류",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "세이웰",
    "status": "거래중"
  },
  {
    "code": "1208548703",
    "name": "온더보더 코엑스 도심공항점 (JRW)",
    "channel": "체인",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "온더보더",
    "status": "거래중"
  },
  {
    "code": "4206400773",
    "name": "세븐패티버거 송리단길점",
    "channel": "체인",
    "manager": "SW",
    "storeType": "매장",
    "brand": "세븐패티버거",
    "status": "거래중"
  },
  {
    "code": "7518500503",
    "name": "온더보더 스타필드하남점 (JRW)",
    "channel": "체인",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "온더보더",
    "status": "거래중"
  },
  {
    "code": "2131641419",
    "name": "세븐패티버거 청량리점",
    "channel": "체인",
    "manager": "SW",
    "storeType": "매장",
    "brand": "세븐패티버거",
    "status": "거래중"
  },
  {
    "code": "5392801821",
    "name": "필앤필버거 김포점",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "필앤필버거 김포점",
    "status": "거래중"
  },
  {
    "code": "6038111772",
    "name": "CJFW(주)부산센터",
    "channel": "기업",
    "manager": "KT",
    "storeType": "비매장",
    "brand": "CJFW",
    "status": "거래중"
  },
  {
    "code": "189-81-00700",
    "name": "주식회사 굿프랜즈",
    "channel": "제조",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "주식회사 굿프랜즈",
    "status": "거래중"
  },
  {
    "code": "1078540256",
    "name": "온더보더 타임스퀘어점 (JRW)",
    "channel": "체인",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "온더보더",
    "status": "거래중"
  },
  {
    "code": "2968500520",
    "name": "온더보더 롯데몰김포공항점 (JRW)",
    "channel": "체인",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "온더보더",
    "status": "거래중"
  },
  {
    "code": "6848502439",
    "name": "후라이드참잘하는집",
    "channel": "체인",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "후라이드참잘하는집",
    "status": "거래중"
  },
  {
    "code": "2102974658",
    "name": "벅벅버거",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "벅벅버거",
    "status": "거래중"
  },
  {
    "code": "2896500617",
    "name": "매드버거",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "매드버거",
    "status": "거래중"
  },
  {
    "code": "3103801103",
    "name": "필앤필버거(PHIL&FILL BURGER)",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "필앤필버거(PHIL&FILL BURGER)",
    "status": "거래중"
  },
  {
    "code": "1877200148",
    "name": "버기즈 탄방점",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "버기즈 탄방점",
    "status": "거래중"
  },
  {
    "code": "3621002813",
    "name": "크레이지버거(가람점)",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "크레이지버거",
    "status": "거래중"
  },
  {
    "code": "5332301911",
    "name": "크레이지버거 밤리단길 일산 밤가시점",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "크레이지버거",
    "status": "거래중"
  },
  {
    "code": "6646700544",
    "name": "버기즈 어은점",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "버기즈 어은점",
    "status": "거래중"
  },
  {
    "code": "3878102559",
    "name": "(주)인맥에프엔씨",
    "channel": "체인",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "(주)인맥에프엔씨",
    "status": "거래중"
  },
  {
    "code": "3363101654",
    "name": "디어버거 진주경상대점",
    "channel": "권역배송",
    "manager": "KT",
    "storeType": "매장",
    "brand": "디어버거",
    "status": "거래중"
  },
  {
    "code": "6838503624",
    "name": "온더보더 하버마스타 (JRW)",
    "channel": "체인",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "온더보더",
    "status": "거래중"
  },
  {
    "code": "3268501475",
    "name": "온더보더 대전테크노중앙로점 (JRW)",
    "channel": "체인",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "온더보더",
    "status": "거래중"
  },
  {
    "code": "4686400794",
    "name": "버거바",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "버거바",
    "status": "거래중"
  },
  {
    "code": "3210404046",
    "name": "크레이지버거 수원 인계점",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "크레이지버거",
    "status": "거래중"
  },
  {
    "code": "5244901144",
    "name": "세븐패티버거 영등포점",
    "channel": "체인",
    "manager": "SW",
    "storeType": "매장",
    "brand": "세븐패티버거",
    "status": "거래중"
  },
  {
    "code": "6498802761",
    "name": "주식회사 로이플렉스",
    "channel": "체인",
    "manager": "NH",
    "storeType": "비매장",
    "brand": "주식회사 로이플렉스",
    "status": "거래중"
  },
  {
    "code": "120-85-52717",
    "name": "제스티살룬 네이버점",
    "channel": "체인",
    "manager": "SW",
    "storeType": "매장",
    "brand": "제스티살룬",
    "status": "거래중"
  },
  {
    "code": "2414100337",
    "name": "킴스키친",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "킴스키친",
    "status": "거래중"
  },
  {
    "code": "1562901851",
    "name": "디어버거_사천점",
    "channel": "권역배송",
    "manager": "KT",
    "storeType": "매장",
    "brand": "디어버거",
    "status": "거래중"
  },
  {
    "code": "7200601274",
    "name": "오일리버거 교동본점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "오일리버거 교동본점",
    "status": "거래중"
  },
  {
    "code": "1835300669",
    "name": "파이어벨 코엑스점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "파이어벨 코엑스점",
    "status": "거래중"
  },
  {
    "code": "3360103190",
    "name": "디어버거 광양점",
    "channel": "권역배송",
    "manager": "KT",
    "storeType": "매장",
    "brand": "디어버거",
    "status": "거래중"
  },
  {
    "code": "999",
    "name": "[온라인] (주)마켓보로",
    "channel": "온라인",
    "manager": "Bomi",
    "storeType": "비매장",
    "brand": "온라인",
    "status": "거래중"
  },
  {
    "code": "8654600647",
    "name": "요리지존(파스타왕)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "요리지존(파스타왕)",
    "status": "거래중"
  },
  {
    "code": "5605100805",
    "name": "디얼버거(DEAR BURGER) (디어버거 진주)",
    "channel": "권역배송",
    "manager": "KT",
    "storeType": "매장",
    "brand": "디어버거",
    "status": "거래중"
  },
  {
    "code": "6294001446",
    "name": "트레인 버거(TRAIN BURGER)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "트레인 버거(TRAIN BURGER)",
    "status": "거래중"
  },
  {
    "code": "6713401621",
    "name": "매드버거 (궁동점)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "매드버거 (궁동점)",
    "status": "거래중"
  },
  {
    "code": "5618101202",
    "name": "주식회사 쉬즈베이글",
    "channel": "체인",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "주식회사 쉬즈베이글",
    "status": "거래중"
  },
  {
    "code": "1288639317",
    "name": "㈜ 엔젤푸드",
    "channel": "도매",
    "manager": "NH",
    "storeType": "비매장",
    "brand": "㈜ 엔젤푸드",
    "status": "거래중"
  },
  {
    "code": "2201006816",
    "name": "제레미버거_양재본점",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "제레미버거_양재본점",
    "status": "거래중"
  },
  {
    "code": "5523901096",
    "name": "뉴욕버거 김해공항점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "뉴욕버거 김해공항점",
    "status": "거래중"
  },
  {
    "code": "1020288010",
    "name": "행루즈버거",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "행루즈버거",
    "status": "거래중"
  },
  {
    "code": "6246100786",
    "name": "1986 시부야버거",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "1986 시부야버거",
    "status": "거래중"
  },
  {
    "code": "4392901722",
    "name": "버기즈 관평점",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "버기즈 관평점",
    "status": "거래중"
  },
  {
    "code": "4570603409",
    "name": "버거고",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "버거고",
    "status": "거래중"
  },
  {
    "code": "4841502234",
    "name": "버거비앤비(광교점)",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "버거비앤비(광교점)",
    "status": "거래중"
  },
  {
    "code": "2081721529",
    "name": "버거타임",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "버거타임",
    "status": "거래중"
  },
  {
    "code": "6031594669",
    "name": "홀리버거(HOLY BURGER)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "홀리버거(HOLY BURGER)",
    "status": "거래중"
  },
  {
    "code": "5500203302",
    "name": "로우로우 버거샵(RAW RAW Burgershop)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "로우로우 버거샵(RAW RAW Burgershop)",
    "status": "거래중"
  },
  {
    "code": "6342801091",
    "name": "피자팬팬",
    "channel": "체인",
    "manager": "NH",
    "storeType": "매장",
    "brand": "피자팬팬",
    "status": "거래중"
  },
  {
    "code": "1787200426",
    "name": "버거비앤비(BURGER BNB)",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "버거비앤비(BURGER BNB)",
    "status": "거래중"
  },
  {
    "code": "6568102756",
    "name": "(주)현대그린푸드",
    "channel": "제조",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "(주)현대그린푸드",
    "status": "거래중"
  },
  {
    "code": "3558503317",
    "name": "온더보더 마곡 코엑스점 (JRW)",
    "channel": "체인",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "온더보더",
    "status": "거래중"
  },
  {
    "code": "2068694885",
    "name": "버거비",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "버거비",
    "status": "거래중"
  },
  {
    "code": "2738502132",
    "name": "온더보더 에버랜드점 (JRW)",
    "channel": "체인",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "온더보더",
    "status": "거래중"
  },
  {
    "code": "6182492568",
    "name": "벅벅버거(이태원점)",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "벅벅버거",
    "status": "거래중"
  },
  {
    "code": "7958801686",
    "name": "(주)테이스테이 (Teistay)_홍대점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "(주)테이스테이 (Teistay)_홍대점",
    "status": "거래중"
  },
  {
    "code": "2880703192",
    "name": "디어버거 도계점",
    "channel": "권역배송",
    "manager": "KT",
    "storeType": "매장",
    "brand": "디어버거",
    "status": "거래중"
  },
  {
    "code": "4938503263",
    "name": "온더보더 롯데프리미엄아울렛 동부산점 (SRW)",
    "channel": "체인",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "온더보더",
    "status": "거래중"
  },
  {
    "code": "8436400277",
    "name": "제레미버거_군자점",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "제레미버거_군자점",
    "status": "거래중"
  },
  {
    "code": "5366600349",
    "name": "제레미버거_선유도점",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "제레미버거_선유도점",
    "status": "거래중"
  },
  {
    "code": "2036403175",
    "name": "백소정 서창점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "백소정",
    "status": "거래중"
  },
  {
    "code": "5101766564",
    "name": "버거쑈",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "버거쑈",
    "status": "거래중"
  },
  {
    "code": "6331202228",
    "name": "백소정 인하대후문점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "백소정",
    "status": "거래중"
  },
  {
    "code": "7283701350",
    "name": "백소정 주안역점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "백소정",
    "status": "거래중"
  },
  {
    "code": "8641702604",
    "name": "신촌버거",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "신촌버거",
    "status": "거래중"
  },
  {
    "code": "3142783173",
    "name": "폼프리츠 카이스트점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "폼프리츠 카이스트점",
    "status": "거래중"
  },
  {
    "code": "1858100902",
    "name": "주식회사 스모크하우스512",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "주식회사 스모크하우스512",
    "status": "거래중"
  },
  {
    "code": "2363800785",
    "name": "써니사이드 다이닝",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "써니사이드 다이닝",
    "status": "거래중"
  },
  {
    "code": "7488602763",
    "name": "주식회사 제스티살룬",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "주식회사 제스티살룬",
    "status": "거래중"
  },
  {
    "code": "3951202156",
    "name": "백소정 서울대입구역점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "백소정",
    "status": "거래중"
  },
  {
    "code": "7958801002",
    "name": "(주)테이스테이_마곡점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "(주)테이스테이_마곡점",
    "status": "거래중"
  },
  {
    "code": "5410200975",
    "name": "소울버킷(SOUL BUCKET)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "소울버킷(SOUL BUCKET)",
    "status": "거래중"
  },
  {
    "code": "3601502660",
    "name": "세븐패티버거 낙성대점",
    "channel": "체인",
    "manager": "SW",
    "storeType": "매장",
    "brand": "세븐패티버거",
    "status": "거래중"
  },
  {
    "code": "7333701314",
    "name": "버거치즈스마일 원곡점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "버거치즈스마일 원곡점",
    "status": "거래중"
  },
  {
    "code": "6384700741",
    "name": "보스턴 김포 장기점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "보스턴 김포 장기점",
    "status": "거래중"
  },
  {
    "code": "5182601650",
    "name": "백소정 옥수점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "백소정",
    "status": "거래중"
  },
  {
    "code": "2346900196",
    "name": "버거401",
    "channel": "권역배송",
    "manager": "KT",
    "storeType": "매장",
    "brand": "버거401",
    "status": "거래중"
  },
  {
    "code": "999-2",
    "name": "[온라인] 쿠팡 고객",
    "channel": "온라인",
    "manager": "bomi",
    "storeType": "비매장",
    "brand": "온라인",
    "status": "거래중"
  },
  {
    "code": "6990103779",
    "name": "떰즈업 (THUMBS UP)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "떰즈업 (THUMBS UP)",
    "status": "거래중"
  },
  {
    "code": "6848500002",
    "name": "후참잘 (케이닥)",
    "channel": "체인",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "후라이드참잘하는집",
    "status": "거래중"
  },
  {
    "code": "2252901370",
    "name": "백소정 과천점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "백소정",
    "status": "거래중"
  },
  {
    "code": "8330202630",
    "name": "백소정 구월로데오점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "백소정",
    "status": "거래중"
  },
  {
    "code": "2050815666",
    "name": "도그즈인번즈(Dogs Buns)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "도그즈인번즈(Dogs Buns)",
    "status": "거래중"
  },
  {
    "code": "4383401494",
    "name": "벅벅버거 (신당점)",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "벅벅버거",
    "status": "거래중"
  },
  {
    "code": "1238573915",
    "name": "온더보더 광명에이케이점 (JRW)",
    "channel": "체인",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "온더보더",
    "status": "거래중"
  },
  {
    "code": "8061202753",
    "name": "버거넛",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "버거넛",
    "status": "거래중"
  },
  {
    "code": "8344800609",
    "name": "힘난다버거 성복점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "힘난다(매장)",
    "status": "거래중"
  },
  {
    "code": "1201210765",
    "name": "버거스하이",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "버거스하이",
    "status": "거래중"
  },
  {
    "code": "2093743295",
    "name": "힘난다버거_광교점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "힘난다(매장)",
    "status": "거래중"
  },
  {
    "code": "3523301567",
    "name": "치코스",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "치코스",
    "status": "거래중"
  },
  {
    "code": "3520403342",
    "name": "백소정 주안아인병원점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "백소정",
    "status": "거래중"
  },
  {
    "code": "5571602276",
    "name": "링키지 버거(LINKAGE BURGER)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "링키지 버거(LINKAGE BURGER)",
    "status": "거래중"
  },
  {
    "code": "6950503135",
    "name": "니즈버거 여의도점",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "니즈버거 여의도점",
    "status": "거래중"
  },
  {
    "code": "7056100751",
    "name": "바이트클럽(BITE CLUB)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "바이트클럽(BITE CLUB)",
    "status": "거래중"
  },
  {
    "code": "3266600698",
    "name": "백소정 역삼GS점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "백소정",
    "status": "거래중"
  },
  {
    "code": "2368701424",
    "name": "엘더버거",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "엘더버거",
    "status": "거래중"
  },
  {
    "code": "8020203407",
    "name": "몽키필리",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "몽키필리",
    "status": "거래중"
  },
  {
    "code": "8276100217",
    "name": "버거그랩(Burger grab)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "버거그랩(Burger grab)",
    "status": "거래중"
  },
  {
    "code": "5024638621",
    "name": "백소정 도곡점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "백소정",
    "status": "거래중"
  },
  {
    "code": "1514200690",
    "name": "포티데이즈",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "포티데이즈",
    "status": "거래중"
  },
  {
    "code": "5111066869",
    "name": "스테이플버거 가산점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "스테이플버거 가산점",
    "status": "거래중"
  },
  {
    "code": "4350803456",
    "name": "라카이(LAKAI)",
    "channel": "권역배송",
    "manager": "KT",
    "storeType": "매장",
    "brand": "라카이(LAKAI)",
    "status": "거래중"
  },
  {
    "code": "6475100586",
    "name": "프레디버거 화곡역점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "프레디버거 화곡역점",
    "status": "거래중"
  },
  {
    "code": "1590502703",
    "name": "백소정 방배역점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "백소정",
    "status": "거래중"
  },
  {
    "code": "2114601043",
    "name": "경성버거",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "경성버거",
    "status": "거래중"
  },
  {
    "code": "8297700325",
    "name": "19버거테이블",
    "channel": "권역배송",
    "manager": "KT",
    "storeType": "매장",
    "brand": "19버거테이블",
    "status": "거래중"
  },
  {
    "code": "7401902132",
    "name": "디어버거 창원대점",
    "channel": "권역배송",
    "manager": "KT",
    "storeType": "매장",
    "brand": "디어버거",
    "status": "거래중"
  },
  {
    "code": "1420673959",
    "name": "온더고(ONTHEGO)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "온더고(ONTHEGO)",
    "status": "거래중"
  },
  {
    "code": "7861402598",
    "name": "노컷서울 금호",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "노컷서울 금호",
    "status": "거래중"
  },
  {
    "code": "5766500734",
    "name": "버거 다이브",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "버거 다이브",
    "status": "거래중"
  },
  {
    "code": "6294200529",
    "name": "꿈버거 상점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "꿈버거 상점",
    "status": "거래중"
  },
  {
    "code": "6780102735",
    "name": "다원푸드",
    "channel": "도매",
    "manager": "NH",
    "storeType": "매장",
    "brand": "다원푸드",
    "status": "거래중"
  },
  {
    "code": "6407800490",
    "name": "니즈버거 을지로점",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "니즈버거 을지로점",
    "status": "거래중"
  },
  {
    "code": "7942200640",
    "name": "든해버거",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "든해버거",
    "status": "거래중"
  },
  {
    "code": "5994400969",
    "name": "백소정 동양미래대점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "백소정",
    "status": "거래중"
  },
  {
    "code": "2445100878",
    "name": "버문스버거 주안시민공원역점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "버문스버거 주안시민공원역점",
    "status": "거래중"
  },
  {
    "code": "1980802072",
    "name": "보스턴수제버거 부평점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "보스턴수제버거 부평점",
    "status": "거래중"
  },
  {
    "code": "7061902199",
    "name": "엉클캐빈",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "엉클캐빈",
    "status": "거래중"
  },
  {
    "code": "1461802390",
    "name": "마시안 300",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "마시안 300",
    "status": "거래중"
  },
  {
    "code": "6490202807",
    "name": "힘난다버거_양평역점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "힘난다(매장)",
    "status": "거래중"
  },
  {
    "code": "4403601486",
    "name": "세포버거",
    "channel": "권역배송",
    "manager": "KT",
    "storeType": "매장",
    "brand": "세포버거",
    "status": "거래중"
  },
  {
    "code": "2301402381",
    "name": "레이지오프(Lazy off)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "레이지오프(Lazy off)",
    "status": "거래중"
  },
  {
    "code": "2631802171",
    "name": "케일럽버거",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "케일럽버거",
    "status": "거래중"
  },
  {
    "code": "5405200337",
    "name": "티제이 버거앤 파스타",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "티제이 버거앤 파스타",
    "status": "거래중"
  },
  {
    "code": "2260672904",
    "name": "버거웍스",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "버거웍스",
    "status": "거래중"
  },
  {
    "code": "2161072511",
    "name": "힘난다버거_상암점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "힘난다(매장)",
    "status": "거래중"
  },
  {
    "code": "5293200875",
    "name": "타운앤컨트리스햄버거스",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "타운앤컨트리스햄버거스",
    "status": "거래중"
  },
  {
    "code": "8195900585",
    "name": "솜다리(솜다리 버거)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "솜다리(솜다리 버거)",
    "status": "거래중"
  },
  {
    "code": "2437700477",
    "name": "조지아미고(georgiamigo)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "조지아미고(georgiamigo)",
    "status": "거래중"
  },
  {
    "code": "2171726339",
    "name": "노컷서울 마곡",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "노컷서울 마곡",
    "status": "거래중"
  },
  {
    "code": "7380203519",
    "name": "조선버거",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "조선버거",
    "status": "거래중"
  },
  {
    "code": "6566400805",
    "name": "오버드라이브",
    "channel": "권역배송",
    "manager": "KT",
    "storeType": "매장",
    "brand": "오버드라이브",
    "status": "거래중"
  },
  {
    "code": "3598502552",
    "name": "레츠잇치킨",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "레츠잇치킨",
    "status": "거래중"
  },
  {
    "code": "7382901726",
    "name": "브라더s",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "브라더s",
    "status": "거래중"
  },
  {
    "code": "8378101910",
    "name": "주식회사 호맥",
    "channel": "제조",
    "manager": "KT",
    "storeType": "비매장",
    "brand": "주식회사 호맥",
    "status": "거래중"
  },
  {
    "code": "8106600761",
    "name": "버거리 인천 루원시티점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "버거리 인천 루원시티점",
    "status": "거래중"
  },
  {
    "code": "4012471719",
    "name": "미분당 영종도점",
    "channel": "체인",
    "manager": "NH",
    "storeType": "매장",
    "brand": "미분당 영종도점",
    "status": "거래중"
  },
  {
    "code": "1164401327",
    "name": "노스트레스버거 한남점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "노스트레스버거 한남점",
    "status": "거래중"
  },
  {
    "code": "5492401932",
    "name": "노스트레스버거 신흥",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "노스트레스버거 신흥",
    "status": "거래중"
  },
  {
    "code": "2890403971",
    "name": "번앤샷 BURN&SHOT",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "번앤샷 BURN&SHOT",
    "status": "거래중"
  },
  {
    "code": "5058528664",
    "name": "프라이도화",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "프라이도화",
    "status": "거래중"
  },
  {
    "code": "6221802287",
    "name": "슬로우키친",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "슬로우키친",
    "status": "거래중"
  },
  {
    "code": "2780702750",
    "name": "비스티버거 반포직영",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "비스티버거 반포직영",
    "status": "거래중"
  },
  {
    "code": "6975300402",
    "name": "브로버거(방이점)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "브로버거(방이점)",
    "status": "거래중"
  },
  {
    "code": "8030101083",
    "name": "더브라이언스",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "더브라이언스",
    "status": "거래중"
  },
  {
    "code": "4291301990",
    "name": "진주 수제버거 LAKAI",
    "channel": "권역배송",
    "manager": "KT",
    "storeType": "매장",
    "brand": "진주 수제버거 LAKAI",
    "status": "거래중"
  },
  {
    "code": "3021192722",
    "name": "성수망치버거",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "성수망치버거",
    "status": "거래중"
  },
  {
    "code": "3393801123",
    "name": "곰버거",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "곰버거",
    "status": "거래중"
  },
  {
    "code": "1981901860",
    "name": "미분당 송내점",
    "channel": "체인",
    "manager": "NH",
    "storeType": "매장",
    "brand": "미분당 송내점",
    "status": "거래중"
  },
  {
    "code": "5054800899",
    "name": "푸라닭 옥수점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "푸라닭 옥수점",
    "status": "거래중"
  },
  {
    "code": "6154700689",
    "name": "푸라닭 정릉1호점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "푸라닭 정릉1호점",
    "status": "거래중"
  },
  {
    "code": "4548502255",
    "name": "콩킨누스",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "콩킨누스",
    "status": "거래중"
  },
  {
    "code": "6481302528",
    "name": "우프스낵바",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "우프스낵바",
    "status": "거래중"
  },
  {
    "code": "3490203037",
    "name": "오엑스 그릴 OXG 도곡점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "오엑스 그릴 OXG 도곡점",
    "status": "거래중"
  },
  {
    "code": "8973001225",
    "name": "그랜드버거",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "그랜드버거",
    "status": "거래중"
  },
  {
    "code": "1772901519",
    "name": "글래디버거(이천중리지구점)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "글래디버거(이천중리지구점)",
    "status": "거래중"
  },
  {
    "code": "5440403091",
    "name": "버거스타디움(BURGERSTADIUM)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "버거스타디움(BURGERSTADIUM)",
    "status": "거래중"
  },
  {
    "code": "2742302025",
    "name": "썬이스트 버거",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "썬이스트 버거",
    "status": "거래중"
  },
  {
    "code": "4691202331",
    "name": "베어스타코 용인성복점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "베어스타코 용인성복점",
    "status": "거래중"
  },
  {
    "code": "3223601835",
    "name": "패티스버거 사당점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "패티스버거 사당점",
    "status": "거래중"
  },
  {
    "code": "3015900861",
    "name": "버거플리즈",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "버거플리즈",
    "status": "거래중"
  },
  {
    "code": "1435700112",
    "name": "싼타클로스(산타버거)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "싼타클로스(산타버거)",
    "status": "거래중"
  },
  {
    "code": "361-85-02609",
    "name": "제스티살룬 목동현대점",
    "channel": "체인",
    "manager": "SW",
    "storeType": "매장",
    "brand": "제스티살룬",
    "status": "거래중"
  },
  {
    "code": "5178802237",
    "name": "(주)제이에스에프엔비",
    "channel": "도매",
    "manager": "NH",
    "storeType": "비매장",
    "brand": "미인계",
    "status": "거래중"
  },
  {
    "code": "7792801604",
    "name": "푸라닭 묵동점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "푸라닭 묵동점",
    "status": "거래중"
  },
  {
    "code": "3325800908",
    "name": "비기버디",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "비기버디",
    "status": "거래중"
  },
  {
    "code": "4505800507",
    "name": "풀리너마이트_홍대본점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "풀리너마이트_홍대본점",
    "status": "거래중"
  },
  {
    "code": "2526700375",
    "name": "니즈(NEEDS)",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "니즈버거 신촌점",
    "status": "거래중"
  },
  {
    "code": "4012354557",
    "name": "레서",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "레서",
    "status": "거래중"
  },
  {
    "code": "7265000941",
    "name": "백소정 약수역점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "백소정",
    "status": "거래중"
  },
  {
    "code": "5080476676",
    "name": "힘난다버거_안동옥동점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "힘난다(매장)",
    "status": "거래중"
  },
  {
    "code": "1213634427",
    "name": "미국버거_봉덕점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "미국버거_봉덕점",
    "status": "거래중"
  },
  {
    "code": "4492002123",
    "name": "프랭크버거 의정부금오신곡점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "프랭크버거 의정부금오신곡점",
    "status": "거래중"
  },
  {
    "code": "1487600307",
    "name": "메이플 버거",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "메이플 버거",
    "status": "거래중"
  },
  {
    "code": "3698702738",
    "name": "주식회사 유안아이앤씨",
    "channel": "도매",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "주식회사 유안아이앤씨",
    "status": "거래중"
  },
  {
    "code": "6848500000",
    "name": "후참잘 (하이닥)",
    "channel": "체인",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "후라이드참잘하는집",
    "status": "거래중"
  },
  {
    "code": "1390282890",
    "name": "은진유통",
    "channel": "도매",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "은진유통",
    "status": "거래중"
  },
  {
    "code": "6845800003",
    "name": "후참잘 (SFN)",
    "channel": "체인",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "후라이드참잘하는집",
    "status": "거래중"
  },
  {
    "code": "6848500001",
    "name": "후참잘 (도한센터)",
    "channel": "체인",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "후라이드참잘하는집",
    "status": "거래중"
  },
  {
    "code": "4442501667",
    "name": "보스턴수제버거",
    "channel": "체인",
    "manager": "NH",
    "storeType": "매장",
    "brand": "보스턴수제버거",
    "status": "거래중"
  },
  {
    "code": "1138106497",
    "name": "(주) 조흥",
    "channel": "도매",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "(주) 조흥",
    "status": "거래중"
  },
  {
    "code": "6921700277",
    "name": "예스한샘유통",
    "channel": "체인물류",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "예스한샘유통",
    "status": "거래중"
  },
  {
    "code": "7438100222",
    "name": "(주)힘난다",
    "channel": "체인",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "힘난다(본사)",
    "status": "거래중"
  },
  {
    "code": "4678601074-1",
    "name": "세이웰_375브런치(아워홈)",
    "channel": "체인물류",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "세이웰",
    "status": "거래중"
  },
  {
    "code": "2103951708",
    "name": "잭스패티",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "잭스패티",
    "status": "거래중"
  },
  {
    "code": "5210303451",
    "name": "치벅",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "치벅",
    "status": "거래중"
  },
  {
    "code": "6595400751",
    "name": "힘난다버거_인천마전점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "힘난다(매장)",
    "status": "거래중"
  },
  {
    "code": "2031312217",
    "name": "미분당 배곧신도시점",
    "channel": "체인",
    "manager": "NH",
    "storeType": "매장",
    "brand": "미분당 배곧신도시점",
    "status": "거래중"
  },
  {
    "code": "522-11-02673",
    "name": "메이상사",
    "channel": "도매",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "이마트 (수출)",
    "status": "거래중"
  },
  {
    "code": "5698603586",
    "name": "주식회사 아란치니브라더스",
    "channel": "제조",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "주식회사 아란치니브라더스",
    "status": "거래중"
  },
  {
    "code": "7398800843",
    "name": "(주)한승인터내쇼날",
    "channel": "도매",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "(주)한승인터내쇼날",
    "status": "거래중"
  },
  {
    "code": "6038111645",
    "name": "CJFW(주)안양센터",
    "channel": "기업",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "CJFW",
    "status": "거래중"
  },
  {
    "code": "121-81-40026",
    "name": "주식회사 한국도매물류",
    "channel": "도매",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "주식회사 한국도매물류",
    "status": "거래중"
  },
  {
    "code": "8031601675",
    "name": "한소쿠리",
    "channel": "도매",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "한소쿠리",
    "status": "거래중"
  },
  {
    "code": "3878102500",
    "name": "(주)인맥-올잇마켓",
    "channel": "체인",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "(주)인맥에프엔씨",
    "status": "거래중"
  },
  {
    "code": "7520203868",
    "name": "이븐버거(EVEN BURGER)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "이븐버거",
    "status": "거래중"
  },
  {
    "code": "5621202038",
    "name": "초가삼간 보스턴수제버거",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "초가삼간 보스턴수제버거",
    "status": "거래중"
  },
  {
    "code": "6138700700",
    "name": "농업회사법인 우리쌀푸드(주)",
    "channel": "제조",
    "manager": "NH",
    "storeType": "비매장",
    "brand": "농업회사법인 우리쌀푸드(주)",
    "status": "거래중"
  },
  {
    "code": "4611602655",
    "name": "치벅(CHEEBUCK) 강남구청점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "치벅(CHEEBUCK) 강남구청점",
    "status": "거래중"
  },
  {
    "code": "765-07-01968",
    "name": "플레이벅",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "플레이벅",
    "status": "거래중"
  },
  {
    "code": "4108700697",
    "name": "주식회사 월광식자재",
    "channel": "식자재마트",
    "manager": "KT",
    "storeType": "비매장",
    "brand": "주식회사 월광식자재",
    "status": "거래중"
  },
  {
    "code": "3082901715",
    "name": "버거플리",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "버거플리",
    "status": "거래중"
  },
  {
    "code": "891881802091",
    "name": "주식회사 에이비랩코리아",
    "channel": "본사",
    "manager": "Bomi",
    "storeType": "비매장",
    "brand": "주식회사 에이비랩코리아",
    "status": "거래중"
  },
  {
    "code": "124-86-61480",
    "name": "가나식품주식회사",
    "channel": "매입",
    "manager": "Bomi",
    "storeType": "비매장",
    "brand": "가나식품주식회사",
    "status": "거래중"
  },
  {
    "code": "에이비랩",
    "name": "직원구매",
    "channel": "본사",
    "manager": "Bomi",
    "storeType": "비매장",
    "brand": "직원구매",
    "status": "거래중"
  },
  {
    "code": "샘플",
    "name": "에이비랩_샘플",
    "channel": "본사",
    "manager": "Bomi",
    "storeType": "비매장",
    "brand": "에이비랩_샘플",
    "status": "거래중"
  },
  {
    "code": "1078112091",
    "name": "제니코식품 (주)",
    "channel": "매입",
    "manager": "Bomi",
    "storeType": "비매장",
    "brand": "제니코식품 (주)",
    "status": "거래중"
  },
  {
    "code": "7862702017",
    "name": "비케이펍앤버거(BKPUP& Burger)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "비케이펍앤버거(BKPUP& Burger)",
    "status": "거래종료"
  },
  {
    "code": "3396400711",
    "name": "세븐패티버거 가산점",
    "channel": "체인",
    "manager": "SW",
    "storeType": "매장",
    "brand": "세븐패티버거",
    "status": "거래종료"
  },
  {
    "code": "5500103678",
    "name": "세븐패티버거 죽전점",
    "channel": "체인",
    "manager": "SW",
    "storeType": "매장",
    "brand": "세븐패티버거",
    "status": "거래종료"
  },
  {
    "code": "748-86-02700",
    "name": "제스티살룬 파주운정점",
    "channel": "체인",
    "manager": "SW",
    "storeType": "매장",
    "brand": "제스티살룬",
    "status": "거래종료"
  },
  {
    "code": "3198701393",
    "name": "주식회사 더백에프앤비",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "주식회사 더백에프앤비",
    "status": "거래종료"
  },
  {
    "code": "3198701302",
    "name": "더백푸드트럭본점(해방촌점)",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "더백푸드트럭본점(해방촌점)",
    "status": "거래종료"
  },
  {
    "code": "6822401113",
    "name": "제레미버거_도산점",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "제레미버거_도산점",
    "status": "거래종료"
  },
  {
    "code": "2688802994",
    "name": "주식회사 오오티오",
    "channel": "체인",
    "manager": "SW",
    "storeType": "매장",
    "brand": "세븐패티버거",
    "status": "거래종료"
  },
  {
    "code": "2580403069",
    "name": "버거타임",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "버거타임",
    "status": "거래종료"
  },
  {
    "code": "6840501122",
    "name": "노컷서울 금호",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "노컷서울 금호",
    "status": "거래종료"
  },
  {
    "code": "8585900862",
    "name": "도그즈인번즈(Dogs Buns)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "도그즈인번즈(Dogs Buns)",
    "status": "거래종료"
  },
  {
    "code": "2073161771",
    "name": "패티패티",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "패티패티",
    "status": "거래종료"
  },
  {
    "code": "4368601065",
    "name": "웰빙푸드시스템 주식회사",
    "channel": "체인",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "666버거",
    "status": "거래종료"
  },
  {
    "code": "5641902066",
    "name": "에이스버거 동대문 본점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "에이스버거 동대문 본점",
    "status": "거래종료"
  },
  {
    "code": "6305200912",
    "name": "브릿지헤드코리아",
    "channel": "도매",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "브릿지헤드코리아",
    "status": "거래종료"
  },
  {
    "code": "1238574777",
    "name": "프레시원주식회사 남서울사업부",
    "channel": "기업",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "CJFW",
    "status": "거래종료"
  },
  {
    "code": "3198722222",
    "name": "주식회사 더백에프앤비(더현대서울점)",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "주식회사 더백에프앤비(더현대서울점)",
    "status": "거래종료"
  },
  {
    "code": "4308503068",
    "name": "테이스티버거 대전둔산점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "테이스티버거 대전둔산점",
    "status": "거래종료"
  },
  {
    "code": "1298688941",
    "name": "(주)에스피씨 지에프에스_피자와썹",
    "channel": "체인",
    "manager": "KT",
    "storeType": "비매장",
    "brand": "SPC(피자와썹)",
    "status": "거래종료"
  },
  {
    "code": "8700902429",
    "name": "바운스무드 오창점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "바운스무드 오창점",
    "status": "거래종료"
  },
  {
    "code": "8631202611",
    "name": "바운스무드 복대점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "바운스무드 복대점",
    "status": "거래종료"
  },
  {
    "code": "2178124157",
    "name": "다인에프씨 주식회사",
    "channel": "도매",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "다인에프씨 주식회사",
    "status": "거래종료"
  },
  {
    "code": "5044114799",
    "name": "유피버거",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "유피버거",
    "status": "거래종료"
  },
  {
    "code": "2378802985-1",
    "name": "주식회사 하라에프에스 (버거옥)",
    "channel": "도매",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "주식회사 하라에프에스 (버거옥)",
    "status": "거래종료"
  },
  {
    "code": "1268639984",
    "name": "농업회사법인 태성그린푸드 주식회사",
    "channel": "체인물류",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "닭장수후라이드",
    "status": "거래종료"
  },
  {
    "code": "1508700999",
    "name": "(주)비스티보이즈(비스티버거)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "(주)비스티보이즈(비스티버거)",
    "status": "거래종료"
  },
  {
    "code": "1750702484",
    "name": "백소정 검단사거리역점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "백소정",
    "status": "거래종료"
  },
  {
    "code": "6524301127",
    "name": "홀리몰리버거_창원",
    "channel": "권역배송",
    "manager": "KT",
    "storeType": "매장",
    "brand": "홀리몰리버거_창원",
    "status": "거래종료"
  },
  {
    "code": "2178124157-1",
    "name": "다인에프씨 주식회사 (다인식품)",
    "channel": "도매",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "다인에프씨 주식회사 (다인식품)",
    "status": "거래종료"
  },
  {
    "code": "7158502918",
    "name": "주식회사 지원글로벌 백소정 계양구청점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "백소정",
    "status": "거래종료"
  },
  {
    "code": "2011014519",
    "name": "더백푸드트럭본점(해방촌점)",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "더백푸드트럭본점(해방촌점)",
    "status": "거래종료"
  },
  {
    "code": "1332002548",
    "name": "포테이토헤드 버거샵",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "포테이토헤드 버거샵",
    "status": "거래종료"
  },
  {
    "code": "2211562422",
    "name": "테이스티버거2018",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "테이스티버거2018",
    "status": "거래종료"
  },
  {
    "code": "7528802086",
    "name": "에이에프컴퍼니",
    "channel": "도매",
    "manager": "NH",
    "storeType": "비매장",
    "brand": "에이에프컴퍼니",
    "status": "거래종료"
  },
  {
    "code": "2964601030",
    "name": "프레클버거(Freckle Burger)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "프레클버거(Freckle Burger)",
    "status": "거래종료"
  },
  {
    "code": "2655501076",
    "name": "파파로카",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "파파로카",
    "status": "거래종료"
  },
  {
    "code": "6473101215",
    "name": "보스턴 영월",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "보스턴 영월",
    "status": "거래종료"
  },
  {
    "code": "7263401009",
    "name": "니즈버거 청라점",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "니즈버거 청라점",
    "status": "거래종료"
  },
  {
    "code": "3437500356",
    "name": "푸라닭 신내점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "푸라닭 신내점",
    "status": "거래종료"
  },
  {
    "code": "8484700668",
    "name": "알지비버거(rgbburger)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "알지비버거(rgbburger)",
    "status": "거래종료"
  },
  {
    "code": "2703201593",
    "name": "홀리몰리버거(HOLYMOLYBURGER)_진해점",
    "channel": "권역배송",
    "manager": "KT",
    "storeType": "매장",
    "brand": "홀리몰리버거(HOLYMOLYBURGER)_진해점",
    "status": "거래종료"
  },
  {
    "code": "3266200685",
    "name": "19버거테이블(부산시청점)",
    "channel": "권역배송",
    "manager": "KT",
    "storeType": "매장",
    "brand": "19버거테이블_부산시청점",
    "status": "거래종료"
  },
  {
    "code": "2791900914",
    "name": "분더버거",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "분더버거",
    "status": "거래종료"
  },
  {
    "code": "1853501369",
    "name": "힘난다버거(코엑스점)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "힘난다버거(코엑스점)",
    "status": "거래종료"
  },
  {
    "code": "8681602633",
    "name": "브로버거(광명점)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "브로버거(광명점)",
    "status": "거래종료"
  },
  {
    "code": "6644700696",
    "name": "버거106(BURGER106)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "버거106",
    "status": "거래종료"
  },
  {
    "code": "7528601875",
    "name": "(주)지인프라임",
    "channel": "체인",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "(주)지인프라임",
    "status": "거래종료"
  },
  {
    "code": "4268702967",
    "name": "백소정 부천역점(㈜다원글로벌)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "백소정",
    "status": "거래종료"
  },
  {
    "code": "1634701208",
    "name": "버거치즈스마일",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "버거치즈스마일",
    "status": "거래종료"
  },
  {
    "code": "6121893990",
    "name": "백소정 철산점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "백소정",
    "status": "거래종료"
  },
  {
    "code": "8938502969",
    "name": "힘난다버거 부천상동점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "힘난다(매장)",
    "status": "거래종료"
  },
  {
    "code": "1028508789",
    "name": "백소정 주안역점(㈜다원글로벌)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "백소정",
    "status": "거래종료"
  },
  {
    "code": "4581801538",
    "name": "미분당 계양구청점",
    "channel": "체인",
    "manager": "NH",
    "storeType": "매장",
    "brand": "미분당 계양구청점",
    "status": "거래종료"
  },
  {
    "code": "7262501026",
    "name": "미분당",
    "channel": "체인",
    "manager": "NH",
    "storeType": "매장",
    "brand": "미분당 청라점",
    "status": "거래종료"
  },
  {
    "code": "4728502796",
    "name": "(주)지원글로벌 백소정역곡역점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "백소정",
    "status": "거래종료"
  },
  {
    "code": "8068603205",
    "name": "(주)양지바른",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "(주)양지바른",
    "status": "거래종료"
  },
  {
    "code": "8795500764",
    "name": "디어버거 창원가로수길점 (본점)",
    "channel": "권역배송",
    "manager": "KT",
    "storeType": "매장",
    "brand": "디어버거",
    "status": "거래종료"
  },
  {
    "code": "3793001014",
    "name": "에프터드링크버거",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "에프터드링크버거",
    "status": "거래종료"
  },
  {
    "code": "2118149636",
    "name": "(주) 제때_서가앤쿡",
    "channel": "도매",
    "manager": "KT",
    "storeType": "비매장",
    "brand": "(주) 제때_서가앤쿡",
    "status": "거래종료"
  },
  {
    "code": "5061711435",
    "name": "오대양푸드",
    "channel": "도매",
    "manager": "KT",
    "storeType": "비매장",
    "brand": "오대양푸드",
    "status": "거래종료"
  },
  {
    "code": "6623401501",
    "name": "디어버거 갤러리아점",
    "channel": "권역배송",
    "manager": "KT",
    "storeType": "매장",
    "brand": "디어버거",
    "status": "거래종료"
  },
  {
    "code": "1160819585",
    "name": "끄트머리피자 (디어버거 사천2호점)",
    "channel": "권역배송",
    "manager": "KT",
    "storeType": "매장",
    "brand": "디어버거",
    "status": "거래종료"
  },
  {
    "code": "1888100196",
    "name": "주식회사 중앙에프엔에스",
    "channel": "식자재마트",
    "manager": "KT",
    "storeType": "비매장",
    "brand": "주식회사 중앙에프엔에스",
    "status": "거래종료"
  },
  {
    "code": "1088176929",
    "name": "주식회사 제이알더블유",
    "channel": "체인",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "온더보더",
    "status": "거래종료"
  },
  {
    "code": "2158679307",
    "name": "유한회사 현아농산",
    "channel": "도매",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "유한회사 현아농산",
    "status": "거래종료"
  },
  {
    "code": "397-86-03531",
    "name": "주식회사 성신",
    "channel": "제조",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "주식회사 성신",
    "status": "거래종료"
  },
  {
    "code": "2757700606",
    "name": "태광유통",
    "channel": "도매",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "태광유통",
    "status": "거래종료"
  },
  {
    "code": "1408170064",
    "name": "주식회사 건우유통",
    "channel": "도매",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "주식회사 건우유통",
    "status": "거래종료"
  },
  {
    "code": "1068514847",
    "name": "해태제과식품 (주) 서울지점",
    "channel": "도매",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "해태제과식품 (주) 서울지점",
    "status": "거래종료"
  },
  {
    "code": "2030431454",
    "name": "코코푸드",
    "channel": "체인물류",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "코코푸드",
    "status": "거래종료"
  },
  {
    "code": "1598603109",
    "name": "주식회사 사우스코어",
    "channel": "제조",
    "manager": "SY",
    "storeType": "비매장",
    "brand": "주식회사 사우스코어",
    "status": "거래종료"
  },
  {
    "code": "3198733333",
    "name": "주식회사 더백에프앤비(작업장)",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "주식회사 더백에프앤비(작업장)",
    "status": "거래종료"
  },
  {
    "code": "3512602100",
    "name": "코코벅",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "코코벅",
    "status": "거래종료"
  },
  {
    "code": "4263900390",
    "name": "에이플렛(A_plat)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "에이플랫(A_PLAT) 버거",
    "status": "거래종료"
  },
  {
    "code": "5383201387",
    "name": "캐비넷베이글",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "캐비넷 베이글",
    "status": "거래종료"
  },
  {
    "code": "7578802325",
    "name": "주식회사 두잇(Doeat)",
    "channel": "제조",
    "manager": "NH",
    "storeType": "비매장",
    "brand": "주식회사 두잇(Doeat)",
    "status": "거래종료"
  },
  {
    "code": "8348502602",
    "name": "주식회사 누땡",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "주식회사 누땡",
    "status": "거래종료"
  },
  {
    "code": "5053510081",
    "name": "백소정 판교점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "백소정",
    "status": "거래종료"
  },
  {
    "code": "2824101088",
    "name": "만나푸드",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "만나푸드",
    "status": "거래종료"
  },
  {
    "code": "2136700797",
    "name": "바운스무드 동남지구점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "바운스무드 동남지구점",
    "status": "거래종료"
  },
  {
    "code": "1683701669",
    "name": "바운스무드 청주 본점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "바운스무드 청주 본점",
    "status": "거래종료"
  },
  {
    "code": "7886400795",
    "name": "아지트카페앤펍(CAFE&PUB)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "아지트카페앤펍(CAFE&PUB)",
    "status": "거래종료"
  },
  {
    "code": "2653300814",
    "name": "스매쉬드 매스",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "스매쉬드 매스",
    "status": "거래종료"
  },
  {
    "code": "1318802877",
    "name": "주식회사 포코너스",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "주식회사 포코너스",
    "status": "거래종료"
  },
  {
    "code": "8150402427",
    "name": "원더샐러드(힘난다버거 코엑스점)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "원더샐러드(힘난다버거 코엑스점)",
    "status": "거래종료"
  },
  {
    "code": "4652201891",
    "name": "아미르 케밥(AMLR KEBAB)",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "아미르 케밥(AMLR KEBAB)",
    "status": "거래종료"
  },
  {
    "code": "0014",
    "name": "브루클린버거더조인트 청계천점",
    "channel": "체인",
    "manager": "NH",
    "storeType": "비매장",
    "brand": "브루클린",
    "status": "거래종료"
  },
  {
    "code": "0018",
    "name": "브루클린더버거조인트 판교점",
    "channel": "체인",
    "manager": "NH",
    "storeType": "비매장",
    "brand": "브루클린",
    "status": "거래종료"
  },
  {
    "code": "0022",
    "name": "브루클린더버거조인트 올림픽공원점",
    "channel": "체인",
    "manager": "NH",
    "storeType": "비매장",
    "brand": "브루클린",
    "status": "거래종료"
  },
  {
    "code": "0010",
    "name": "브루클린더버거조인트 여의도점",
    "channel": "체인",
    "manager": "NH",
    "storeType": "비매장",
    "brand": "브루클린",
    "status": "거래종료"
  },
  {
    "code": "0013",
    "name": "브루클린더버거조인트 분당정자점",
    "channel": "체인",
    "manager": "NH",
    "storeType": "비매장",
    "brand": "브루클린",
    "status": "거래종료"
  },
  {
    "code": "0016",
    "name": "브루클린더버거조인트 목동점",
    "channel": "체인",
    "manager": "NH",
    "storeType": "비매장",
    "brand": "브루클린",
    "status": "거래종료"
  },
  {
    "code": "0011",
    "name": "브루클린더버거조인트 롯데월드몰점",
    "channel": "체인",
    "manager": "NH",
    "storeType": "비매장",
    "brand": "브루클린",
    "status": "거래종료"
  },
  {
    "code": "0009",
    "name": "브루클린더버거조인트 동부이촌점",
    "channel": "체인",
    "manager": "NH",
    "storeType": "비매장",
    "brand": "브루클린",
    "status": "거래종료"
  },
  {
    "code": "0023",
    "name": "브루클린 더 버거 조인트 역삼점",
    "channel": "체인",
    "manager": "NH",
    "storeType": "비매장",
    "brand": "브루클린",
    "status": "거래종료"
  },
  {
    "code": "5868503003",
    "name": "(주)임어전씨 수제버거 전곡역지점",
    "channel": "체인",
    "manager": "NH",
    "storeType": "매장",
    "brand": "(주)임어전씨 수제버거 전곡역지점",
    "status": "거래종료"
  },
  {
    "code": "0003",
    "name": "(주)에스씨비에이치 서래마을점",
    "channel": "체인",
    "manager": "NH",
    "storeType": "비매장",
    "brand": "브루클린",
    "status": "거래종료"
  },
  {
    "code": "0004",
    "name": "(주)에스씨비에이치 삼성점",
    "channel": "체인",
    "manager": "NH",
    "storeType": "비매장",
    "brand": "브루클린",
    "status": "거래종료"
  },
  {
    "code": "0008",
    "name": "(주)에스씨비에이치 광화문 디타워점",
    "channel": "체인",
    "manager": "NH",
    "storeType": "비매장",
    "brand": "브루클린",
    "status": "거래종료"
  },
  {
    "code": "0002",
    "name": "(주)에스씨비에이치 가로수길점",
    "channel": "체인",
    "manager": "NH",
    "storeType": "비매장",
    "brand": "브루클린",
    "status": "거래종료"
  },
  {
    "code": "3134500718",
    "name": "페어쉬림프세라믹",
    "channel": "도매",
    "manager": "NH",
    "storeType": "비매장",
    "brand": "페어쉬림프세라믹",
    "status": "거래종료"
  },
  {
    "code": "5900503179",
    "name": "힘난다버거 송도점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "힘난다(매장)",
    "status": "거래종료"
  },
  {
    "code": "3061451272",
    "name": "피자탑 포항남구점(블럭키친 포항남구점)",
    "channel": "권역배송",
    "manager": "KT",
    "storeType": "매장",
    "brand": "블럭키친",
    "status": "거래종료"
  },
  {
    "code": "1691901106",
    "name": "미스터키친 (블럭키친 경남고성점)",
    "channel": "권역배송",
    "manager": "KT",
    "storeType": "매장",
    "brand": "블럭키친",
    "status": "거래종료"
  },
  {
    "code": "6108702830",
    "name": "주식회사 이음에프엔아이",
    "channel": "도매",
    "manager": "KT",
    "storeType": "비매장",
    "brand": "주식회사 이음에프엔아이",
    "status": "거래종료"
  },
  {
    "code": "1063301362",
    "name": "제일유통",
    "channel": "도매",
    "manager": "KT",
    "storeType": "비매장",
    "brand": "제일유통",
    "status": "거래종료"
  },
  {
    "code": "3948702433",
    "name": "이음푸드",
    "channel": "도매",
    "manager": "KT",
    "storeType": "비매장",
    "brand": "이음푸드",
    "status": "거래종료"
  },
  {
    "code": "3128125280",
    "name": "(주)동원홈푸드",
    "channel": "도매",
    "manager": "KT",
    "storeType": "비매장",
    "brand": "(주)동원홈푸드",
    "status": "거래종료"
  },
  {
    "code": "4555101045",
    "name": "선산맛닭",
    "channel": "권역배송",
    "manager": "KT",
    "storeType": "매장",
    "brand": "선산맛닭",
    "status": "거래종료"
  },
  {
    "code": "3918502857",
    "name": "프레시원주식회사 부산사업부",
    "channel": "기업",
    "manager": "KT",
    "storeType": "비매장",
    "brand": "CJFW",
    "status": "거래종료"
  },
  {
    "code": "7291502248",
    "name": "조선버거",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "조선버거",
    "status": "거래종료"
  },
  {
    "code": "4821501961",
    "name": "행루즈버거",
    "channel": "권역배송",
    "manager": "SW",
    "storeType": "매장",
    "brand": "행루즈버거",
    "status": "거래종료"
  },
  {
    "code": "7320402743",
    "name": "테이스티버거 대전둔산점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "테이스티버거 대전둔산점",
    "status": "거래종료"
  },
  {
    "code": "5935700879",
    "name": "버거플리",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "버거플리",
    "status": "거래종료"
  },
  {
    "code": "5221202516",
    "name": "백소정 서창점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "백소정",
    "status": "거래종료"
  },
  {
    "code": "1051996953",
    "name": "버거넛",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "버거넛",
    "status": "거래종료"
  },
  {
    "code": "4663501287",
    "name": "백소정 도곡점",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "백소정",
    "status": "거래종료"
  },
  {
    "code": "6371902253",
    "name": "오버드라이브",
    "channel": "권역배송",
    "manager": "KT",
    "storeType": "매장",
    "brand": "오버드라이브",
    "status": "거래종료"
  },
  {
    "code": "1772601916",
    "name": "버거바",
    "channel": "권역배송",
    "manager": "NH",
    "storeType": "매장",
    "brand": "버거바",
    "status": "거래종료"
  }
];

const initialTargets: TargetRecord[] = [
  { storeType: "매장", month: thisMonth(), amount: 20000000 },
  { storeType: "비매장", month: thisMonth(), amount: 15000000 },
];

const initialEsts: EstRecord[] = [
  { storeCode: "A001", storeName: "강남점", month: thisMonth(), amount: 18000000 },
  { storeCode: "A002", storeName: "홍대점", month: thisMonth(), amount: 14000000 },
];

const initialSales: SalesRecord[] = [
  makeSale("current", thisMonth(), `${thisMonth()}-17`, "A001", "강남점", "P001", "샘플상품", 1, 12500000, 8000000, 4500000, initialStores),
  makeSale("current", thisMonth(), `${thisMonth()}-17`, "A002", "홍대점", "P001", "샘플상품", 1, 5300000, 3500000, 1800000, initialStores),
];

const initialTimeConfigs: TimeConfig[] = [{ month: thisMonth(), holidays: [] }];

function today() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

function thisMonth() {
  return today().slice(0, 7);
}

function monthStart(month: string) {
  return `${month}-01`;
}

function monthEnd(month: string) {
  const [y, m] = month.split("-").map(Number);
  return `${month}-${String(new Date(y, m, 0).getDate()).padStart(2, "0")}`;
}

function parseYmd(date: string) {
  const [y, m, d] = date.split("-").map(Number);
  return { y, m, d };
}

function addDays(date: string, days: number) {
  const { y, m, d } = parseYmd(date);
  const next = new Date(y, m - 1, d + days);
  return `${next.getFullYear()}-${String(next.getMonth() + 1).padStart(2, "0")}-${String(next.getDate()).padStart(2, "0")}`;
}

function previousMonth(month: string) {
  const [y, m] = month.split("-").map(Number);
  const d = new Date(y, m - 2, 1);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

function addMonths(month: string, diff: number) {
  const [y, m] = month.split("-").map(Number);
  const d = new Date(y, m - 1 + diff, 1);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

function threeMonthStart(month: string) {
  return monthStart(addMonths(month, -2));
}

function sortArrow(active: boolean, direction: SortDirection) {
  if (!active) return "↕";
  return direction === "asc" ? "↑" : "↓";
}

function previousYearMonth(month: string) {
  const [y, m] = month.split("-").map(Number);
  return `${y - 1}-${String(m).padStart(2, "0")}`;
}

function sameDayPrevMonth(date: string) {
  const [y, m, day] = date.split("-").map(Number);
  const d = new Date(y, m - 2, 1);
  const last = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(Math.min(day, last)).padStart(2, "0")}`;
}

function sameDayPrevYear(date: string) {
  const [y, m, day] = date.split("-").map(Number);
  const last = new Date(y - 1, m, 0).getDate();
  return `${y - 1}-${String(m).padStart(2, "0")}-${String(Math.min(day, last)).padStart(2, "0")}`;
}

function inRange(date: string, start: string, end: string) {
  return date >= start && date <= end;
}

function daysBetween(fromDate: string, toDate: string) {
  if (!fromDate || fromDate === "-") return 9999;
  const from = new Date(`${fromDate}T00:00:00`);
  const to = new Date(`${toDate}T00:00:00`);
  const diff = Math.floor((to.getTime() - from.getTime()) / (1000 * 60 * 60 * 24));
  return Number.isFinite(diff) ? diff : 9999;
}

function won(n: number) {
  return Math.round(n || 0).toLocaleString("ko-KR");
}

function pct(n: number) {
  return Number.isFinite(n) ? `${n.toFixed(1)}%` : "-";
}

function num(v: unknown) {
  if (typeof v === "number") return Number.isFinite(v) ? v : 0;

  const text = String(v ?? "").trim();
  if (!text) return 0;

  const isParenthesesNegative = /^\(.*\)$/.test(text);
  const cleaned = text
    .replace(/\u2212/g, "-")
    .replace(/,/g, "")
    .replace(/%/g, "")
    .replace(/원/g, "")
    .replace(/\s/g, "")
    .replace(/[()]/g, "");

  const parsed = Number(cleaned);
  if (!Number.isFinite(parsed)) return 0;
  return isParenthesesNegative ? -Math.abs(parsed) : parsed;
}

function norm(v: unknown) {
  return String(v ?? "").trim();
}

function normalizeStatus(v: unknown): Store["status"] {
  const t = norm(v);
  return t === "종료" || t === "거래종료" || t === "비활성" || t === "비활성화" || t.toLowerCase() === "inactive" || t.toLowerCase() === "closed" ? "거래종료" : "거래중";
}

function normalizeChannel(v: unknown) {
  return norm(v) || "미지정";
}

function normalizeStoreType(v: unknown, fallbackChannel?: unknown) {
  const t = norm(v);
  if (t) return t;
  const ch = norm(fallbackChannel);
  return ch === "매장" ? "매장" : "비매장";
}


function dateText(v: unknown) {
  if (v instanceof Date) {
    return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, "0")}-${String(v.getDate()).padStart(2, "0")}`;
  }
  if (typeof v === "number") {
    const d = XLSX.SSF.parse_date_code(v);
    if (d) return `${d.y}-${String(d.m).padStart(2, "0")}-${String(d.d).padStart(2, "0")}`;
  }
  const t = norm(v).replaceAll(".", "-").replaceAll("/", "-");
  if (/^\d{8}$/.test(t)) return `${t.slice(0, 4)}-${t.slice(4, 6)}-${t.slice(6, 8)}`;
  if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(t)) {
    const [y, m, d] = t.split("-");
    return `${y}-${m.padStart(2, "0")}-${d.padStart(2, "0")}`;
  }
  return t;
}

function monthText(v: unknown) {
  const t = norm(v).replaceAll(".", "-").replaceAll("/", "-");
  if (/^\d{6}$/.test(t)) return `${t.slice(0, 4)}-${t.slice(4, 6)}`;
  if (/^\d{4}-\d{1,2}/.test(t)) {
    const [y, m] = t.split("-");
    return `${y}-${m.padStart(2, "0")}`;
  }
  return t;
}

function dayWeight(date: string, holidays: string[]) {
  const { y, m, d } = parseYmd(date);
  const day = new Date(y, m - 1, d).getDay();
  const isHoliday = holidays.includes(date);
  if (day === 0) return 0;
  if (day === 6) return 0.5;
  if (isHoliday) return 0.5;
  return 1;
}

function getTimeGone(month: string, date: string, timeConfigs: TimeConfig[]) {
  const config = timeConfigs.find((c) => c.month === month);
  const holidays = config?.holidays || [];
  const start = monthStart(month);
  const end = monthEnd(month);
  let totalDays = 0;
  let progressedDays = 0;

  for (let d = start, guard = 0; d <= end && guard < 40; d = addDays(d, 1), guard += 1) {
    const w = dayWeight(d, holidays);
    totalDays += w;
    if (d <= date) progressedDays += w;
  }

  const remainingDays = Math.max(totalDays - progressedDays, 0);
  const timeGoneRate = totalDays ? (progressedDays / totalDays) * 100 : 0;
  return { totalDays, progressedDays, remainingDays, timeGoneRate };
}

function storeMap(stores: Store[]) {
  return new Map(stores.map((s) => [s.code, s]));
}

function normalizeStoreNameKey(value: unknown) {
  return norm(value)
    .toLowerCase()
    .replace(/\s/g, "")
    .replace(/주식회사/g, "")
    .replace(/\(주\)/g, "")
    .replace(/㈜/g, "")
    .replace(/[._-]/g, "");
}

function resolveStoreInfo(storeCode: string, fallbackName: string, fallback: Partial<Store>, stores: Store[]) {
  const original = storeMap(stores).get(storeCode);

  return {
    code: storeCode || fallbackName || "미지정",
    name: original?.name || fallbackName || storeCode || "미지정",
    channel: original?.channel || fallback.channel || "미지정",
    manager: original?.manager || fallback.manager || "미지정",
    storeType: original?.storeType || fallback.storeType || "비매장",
    brand: original?.brand || fallback.brand || "미지정",
    status: original?.status || "거래중",
    originalCode: storeCode,
    originalName: original?.name || fallbackName || storeCode,
  };
}

function makeSale(
  period: PeriodType,
  refMonth: string,
  saleDate: string,
  storeCode: string,
  storeName: string,
  itemCode: string,
  itemName: string,
  quantity: number,
  salesAmount: number,
  costAmount: number,
  profitAmount: number,
  stores: Store[],
  uploadedProfitRate?: number
): SalesRecord {
  const s = storeMap(stores).get(storeCode);
  const profitRate = Number.isFinite(uploadedProfitRate) ? Number(uploadedProfitRate) : salesAmount ? (profitAmount / salesAmount) * 100 : 0;
  return {
    id: `${period}|${refMonth}|${saleDate}|${storeCode}|${itemCode}|${itemName}`,
    period,
    refMonth,
    saleDate,
    storeCode,
    storeName: s?.name || storeName || storeCode,
    channel: s?.channel || "매장",
    manager: s?.manager || "",
    storeType: s?.storeType || "매장",
    brand: s?.brand || "미지정",
    itemCode,
    itemName,
    quantity,
    salesAmount,
    costAmount,
    profitAmount,
    profitRate,
  };
}

type AppStateRow<T> = {
  id: string;
  data: T;
  updated_at?: string;
};

function supabaseConfig() {
  const url = process.env.NEXT_PUBLIC_SUPABASE_URL?.trim().replace(/\/$/, "");
  const key = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY?.trim();

  if (!url || !key) return null;

  return {
    endpoint: `${url}/rest/v1/app_state`,
    headers: {
      apikey: key,
      Authorization: `Bearer ${key}`,
      "Content-Type": "application/json",
    },
  };
}

async function loadSharedState<T>(key: string): Promise<T | null> {
  const config = supabaseConfig();
  if (!config) return null;

  const response = await fetch(`${config.endpoint}?id=eq.${encodeURIComponent(key)}&select=data&limit=1`, {
    method: "GET",
    headers: config.headers,
    cache: "no-store",
  });

  if (!response.ok) throw new Error(`Supabase load failed: ${response.status}`);
  const rows = (await response.json()) as AppStateRow<T>[];
  return rows[0]?.data ?? null;
}

function localMetaKey(key: string) {
  return `${key}__local_meta`;
}

function getLocalMeta(key: string) {
  if (typeof window === "undefined") return { editedAt: 0, pending: false };
  try {
    const parsed = JSON.parse(window.localStorage.getItem(localMetaKey(key)) || "{}");
    return {
      editedAt: Number(parsed.editedAt || 0),
      pending: Boolean(parsed.pending),
    };
  } catch {
    return { editedAt: 0, pending: false };
  }
}

function setLocalMeta(key: string, meta: { editedAt: number; pending: boolean }) {
  if (typeof window === "undefined") return;
  window.localStorage.setItem(localMetaKey(key), JSON.stringify(meta));
}

async function saveSharedState<T>(key: string, value: T) {
  const config = supabaseConfig();
  if (!config) return;

  const response = await fetch(`${config.endpoint}?on_conflict=id`, {
    method: "POST",
    headers: {
      ...config.headers,
      Prefer: "resolution=merge-duplicates,return=minimal",
    },
    body: JSON.stringify({
      id: key,
      data: value,
      updated_at: new Date().toISOString(),
    }),
  });

  if (!response.ok) throw new Error(`Supabase save failed: ${response.status}`);
}

function useLocal<T>(key: string, initial: T) {
  const [value, rawSetValue] = useState<T>(initial);
  const [loaded, setLoaded] = useState(false);
  const valueRef = useRef<T>(initial);
  const saveTimerRef = useRef<number | null>(null);
  const lastSavedJsonRef = useRef("");
  const keyRef = useRef(key);

  useEffect(() => {
    keyRef.current = key;
  }, [key]);

  const persistNow = (nextValue: T) => {
    if (typeof window === "undefined") return;

    const json = JSON.stringify(nextValue);
    const editedAt = Date.now();
    valueRef.current = nextValue;
    lastSavedJsonRef.current = json;
    window.localStorage.setItem(keyRef.current, json);
    setLocalMeta(keyRef.current, { editedAt, pending: true });

    if (saveTimerRef.current) window.clearTimeout(saveTimerRef.current);
    saveTimerRef.current = window.setTimeout(() => {
      saveSharedState(keyRef.current, nextValue)
        .then(() => {
          const meta = getLocalMeta(keyRef.current);
          if (meta.editedAt === editedAt) setLocalMeta(keyRef.current, { editedAt, pending: false });
        })
        .catch((error) => {
          // Supabase 저장이 실패해도 화면의 최신 업로드값은 유지합니다.
          // 다음 업로드/수정 때 다시 저장을 시도합니다.
          console.warn("공유 데이터 저장 실패, 브라우저 저장소에 최신 데이터를 유지합니다.", error);
          const meta = getLocalMeta(keyRef.current);
          if (meta.editedAt === editedAt) setLocalMeta(keyRef.current, { editedAt, pending: true });
        });
    }, 120);
  };

  const setValue: React.Dispatch<React.SetStateAction<T>> = (next) => {
    rawSetValue((prev) => {
      const resolved = typeof next === "function" ? (next as (prev: T) => T)(prev) : next;
      persistNow(resolved);
      return resolved;
    });
  };

  useEffect(() => {
    let cancelled = false;

    async function hydrate() {
      try {
        const localSaved = window.localStorage.getItem(key);
        const meta = getLocalMeta(key);

        if (localSaved && !cancelled) {
          const parsed = JSON.parse(localSaved) as T;
          valueRef.current = parsed;
          lastSavedJsonRef.current = JSON.stringify(parsed);
          rawSetValue(parsed);
        }

        // 로컬에 업로드/수정 데이터가 있으면 원격의 과거값으로 덮어쓰지 않습니다.
        // 로컬이 비어있는 첫 접속일 때만 Supabase 값을 불러옵니다.
        if (!localSaved) {
          const remoteSaved = await loadSharedState<T>(key);
          if (!cancelled && remoteSaved !== null) {
            const remoteJson = JSON.stringify(remoteSaved);
            valueRef.current = remoteSaved;
            lastSavedJsonRef.current = remoteJson;
            window.localStorage.setItem(key, remoteJson);
            setLocalMeta(key, { editedAt: meta.editedAt || Date.now(), pending: false });
            rawSetValue(remoteSaved);
          }
        }
      } catch (error) {
        console.warn("공유 데이터 불러오기 실패, 브라우저 저장소를 우선 사용합니다.", error);
      } finally {
        if (!cancelled) setLoaded(true);
      }
    }

    hydrate();

    return () => {
      cancelled = true;
      if (saveTimerRef.current) window.clearTimeout(saveTimerRef.current);
    };
  }, [key]);

  useEffect(() => {
    if (!loaded) return;

    const retryPendingSave = async () => {
      try {
        const meta = getLocalMeta(key);
        if (!meta.pending) return;
        const localSaved = window.localStorage.getItem(key);
        if (!localSaved) return;
        const parsed = JSON.parse(localSaved) as T;
        await saveSharedState(key, parsed);
        setLocalMeta(key, { editedAt: meta.editedAt, pending: false });
      } catch {
        // 저장 재시도 실패 시에도 화면 데이터는 유지합니다.
      }
    };

    const interval = window.setInterval(retryPendingSave, 5000);
    window.addEventListener("focus", retryPendingSave);

    return () => {
      window.clearInterval(interval);
      window.removeEventListener("focus", retryPendingSave);
    };
  }, [key, loaded]);

  return [value, setValue] as const;
}

function readFileRows(file: File): Promise<Record<string, unknown>[]> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: "array", cellDates: true });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "" });
        const headerIndex = matrix.findIndex((row) =>
          row.some((cell) => ["거래처코드", "거래처명", "주문일", "판매일", "매출일", "Target", "EST"].includes(norm(cell)))
        );

        if (headerIndex < 0) {
          resolve(XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: "" }));
          return;
        }

        const headers = matrix[headerIndex].map((h) => norm(h));
        const rows = matrix.slice(headerIndex + 1)
          .filter((row) => row.some((cell) => norm(cell)))
          .map((row) => {
            const item: Record<string, unknown> = {};
            headers.forEach((header, index) => {
              if (header) item[header] = row[index] ?? "";
            });
            return item;
          });

        resolve(rows);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function groupByKey(records: SalesRecord[], keyFn: (r: SalesRecord) => string) {
  const map = new Map<string, SalesRecord[]>();
  records.forEach((r) => {
    const key = keyFn(r) || "미지정";
    map.set(key, [...(map.get(key) || []), r]);
  });
  return map;
}

function sum(records: SalesRecord[], key: keyof Pick<SalesRecord, "salesAmount" | "costAmount" | "profitAmount" | "quantity">) {
  return records.reduce((a, b) => a + Number(b[key] || 0), 0);
}

function profitRateValue(value: unknown) {
  if (value === undefined || value === null || norm(value) === "") return undefined;

  if (typeof value === "number") {
    return Math.abs(value) > 0 && Math.abs(value) <= 1 ? value * 100 : value;
  }

  const text = norm(value);
  const parsed = num(text);
  return text.includes("%") ? parsed : Math.abs(parsed) > 0 && Math.abs(parsed) <= 1 ? parsed * 100 : parsed;
}

function weightedProfitRate(records: SalesRecord[]) {
  const salesTotal = sum(records, "salesAmount");
  if (!salesTotal) return 0;

  const weighted = records.reduce((total, row) => total + Number(row.salesAmount || 0) * Number(row.profitRate || 0), 0);
  return weighted / salesTotal;
}

function exportExcel(rows: Record<string, string | number>[], fileName: string) {
  if (!rows.length) {
    alert("다운로드할 데이터가 없습니다.");
    return;
  }

  const ws = XLSX.utils.json_to_sheet(rows);
  ws["!cols"] = Object.keys(rows[0]).map((key) => ({
    wch: Math.max(12, key.length + 4, ...rows.map((row) => String(row[key] ?? "").length + 2)),
  }));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  XLSX.writeFile(wb, `${fileName}.xlsx`);
}

function drillPeriodLabel(period: DrillPeriod) {
  if (period === "prevYear") return "전년동월 매출";
  if (period === "prevMonth") return "전월 매출";
  if (period === "current") return "당일까지의 매출";
  return "당월 발주 총 금액";
}

function orderRowsForExcel(rows: SalesRecord[]) {
  return rows.map((r) => ({
    주문일: r.saleDate,
    기준월: r.refMonth,
    기간구분: r.period === "current" ? "당월" : r.period === "prevMonth" ? "전월" : "전년동월",
    거래처코드: r.storeCode,
    거래처명: r.storeName,
    브랜드: r.brand,
    담당자: r.manager || "미지정",
    채널: r.channel,
    상품코드: r.itemCode,
    상품명: r.itemName,
    수량: r.quantity,
    매출금액: r.salesAmount,
    원가금액: r.costAmount,
    이익금액: r.profitAmount,
    이익률: pct(r.profitRate),
  }));
}

const ADMIN_PASSWORD = "ablab2026";

export default function SalesReportClient() {
  const [active, setActive] = useState("대시보드");
  const [isAdmin, setIsAdmin] = useState(false);
  const [stores, setStores] = useLocal<Store[]>("ablab_stores_v15", initialStores);
  const [sales, setSales] = useLocal<SalesRecord[]>("ablab_sales_v14", initialSales);
  const [targets, setTargets] = useLocal<TargetRecord[]>("ablab_targets_v14", initialTargets);
  const [ests, setEsts] = useLocal<EstRecord[]>("ablab_ests_v14", initialEsts);
  const [timeConfigs, setTimeConfigs] = useLocal<TimeConfig[]>("ablab_time_configs_v14", initialTimeConfigs);
  const [codeMappings, setCodeMappings] = useLocal<StoreCodeMapping[]>("ablab_code_mappings_v1", []);
  const [dashMonth, setDashMonth] = useState(thisMonth());
  const [dashDate, setDashDate] = useState(today());

  useEffect(() => {
    if (!dashDate.startsWith(dashMonth)) setDashDate(monthEnd(dashMonth));
  }, [dashMonth, dashDate]);

  useEffect(() => {
    if (!isAdmin && active === "월초관리") setActive("대시보드");
  }, [isAdmin, active]);

  const tg = useMemo(() => getTimeGone(dashMonth, dashDate, timeConfigs), [dashMonth, dashDate, timeConfigs]);
  const menus = isAdmin ? ["대시보드", "매출현황", "품목분석", "월초관리"] : ["대시보드", "매출현황", "품목분석"];

  function adminLogin() {
    const password = window.prompt("관리자 비밀번호를 입력하세요.");
    if (password === ADMIN_PASSWORD) {
      setIsAdmin(true);
      alert("관리자 모드로 전환되었습니다.");
    } else if (password !== null) {
      alert("비밀번호가 올바르지 않습니다.");
    }
  }

  return (
    <main className="flex min-h-screen bg-white text-slate-900" style={{ fontFamily: '"Malgun Gothic", "맑은 고딕", sans-serif' }}>
      <aside className="flex min-h-screen w-44 shrink-0 flex-col border-r border-slate-200 bg-white text-slate-900">
        <div className="border-b border-slate-200 bg-orange-50 p-4 text-base font-bold tracking-tight text-orange-950">에이비랩 코리아 Sales Report</div>
        <nav className="space-y-2 p-3">
          {menus.map((m, index) => (
            <button
              key={m}
              onClick={() => setActive(m)}
              className={`w-full rounded-xl px-3 py-2.5 text-left text-xs font-semibold transition ${
                active === m ? "bg-orange-500 text-white shadow" : "text-slate-700 hover:bg-orange-100 hover:text-orange-900"
              }`}
            >
              {index + 1}. {m}
            </button>
          ))}
        </nav>
        <div className="mt-auto border-t border-slate-200 p-3">
          {isAdmin ? (
            <button
              type="button"
              onClick={() => { setIsAdmin(false); setActive("대시보드"); }}
              className="w-full rounded-xl bg-orange-100 px-3 py-2 text-left text-xs font-semibold text-orange-900 hover:bg-orange-200"
            >
              관리자 모드 해제
            </button>
          ) : (
            <button
              type="button"
              onClick={adminLogin}
              className="w-full rounded-xl bg-orange-100 px-3 py-2 text-left text-xs font-semibold text-orange-900 hover:bg-orange-200"
            >
              관리자 로그인
            </button>
          )}
        </div>
      </aside>

      <section className="min-w-0 flex-1 p-4 lg:p-5">
        <div className="mb-4 rounded-2xl border border-slate-200/70 bg-white/80 p-4 shadow-sm backdrop-blur">
          <div className="flex flex-col gap-4 xl:flex-row xl:items-center xl:justify-between">
            <div>
              <h1 className="text-2xl font-bold tracking-tight">에이비랩 코리아 Sales Report</h1>
            </div>

            <div className="rounded-2xl border border-slate-200/70 bg-slate-50/75 p-3 shadow-sm backdrop-blur">
              <div className="flex flex-wrap items-end gap-2">
                <label className="w-[135px] text-[12px] font-semibold text-slate-600">
                  기준년월
                  <input
                    type="month"
                    value={dashMonth}
                    onChange={(e) => setDashMonth(e.target.value)}
                    className="mt-1 w-full rounded-lg border border-slate-300 bg-white/80 px-2 py-1.5 text-sm font-semibold outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                  />
                </label>
                <label className="w-[145px] text-[12px] font-semibold text-slate-600">
                  기준일
                  <input
                    type="date"
                    value={dashDate}
                    min={monthStart(dashMonth)}
                    max={monthEnd(dashMonth)}
                    onChange={(e) => setDashDate(e.target.value)}
                    className="mt-1 w-full rounded-lg border border-slate-300 bg-white/80 px-2 py-1.5 text-sm font-semibold outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                  />
                </label>
                <HeaderTimeInfo title="총일수" value={tg.totalDays} />
                <HeaderTimeInfo title="진행일수" value={tg.progressedDays} />
                <HeaderTimeInfo title="잔여일수" value={tg.remainingDays} />
              </div>
            </div>
          </div>
        </div>

        {active === "대시보드" && (
          <Dashboard
            stores={stores}
            sales={sales}
            targets={targets}
            ests={ests}
            month={dashMonth}
            date={dashDate}
            timeGone={tg}
            codeMappings={codeMappings}
          />
        )}
        {active === "매출현황" && (
          <SalesStatus
            stores={stores}
            sales={sales}
            targets={targets}
            ests={ests}
            month={dashMonth}
            date={dashDate}
            timeGone={tg}
            codeMappings={codeMappings}
          />
        )}
        {active === "품목분석" && (
          <ItemAnalysis
            stores={stores}
            sales={sales}
            month={dashMonth}
            date={dashDate}
          />
        )}
        {isAdmin && active === "월초관리" && (
          <MonthStartManagement
            stores={stores}
            setStores={setStores}
            sales={sales}
            setSales={setSales}
            targets={targets}
            setTargets={setTargets}
            ests={ests}
            setEsts={setEsts}
            month={dashMonth}
            date={dashDate}
            timeConfigs={timeConfigs}
            setTimeConfigs={setTimeConfigs}
            codeMappings={codeMappings}
            setCodeMappings={setCodeMappings}
          />
        )}
      </section>
    </main>
  );
}

function HeaderTimeInfo({ title, value }: { title: string; value: number }) {
  return (
    <div className="min-w-[76px] rounded-lg border border-slate-200/70 bg-white/70 px-3 py-1.5 text-center shadow-sm">
      <p className="text-[11px] font-semibold text-slate-500">{title}</p>
      <p className="mt-0.5 text-[16px] font-bold text-slate-900">{Number.isInteger(value) ? value : value.toFixed(1)}일</p>
    </div>
  );
}

function KpiGroup({ items }: { items: { title: string; value: string | number; color?: string; format?: "won" | "percent" | "number" }[] }) {
  return (
    <div className="h-full rounded-xl border border-slate-200/70 bg-slate-50/75 p-3 shadow-sm backdrop-blur">
      <div className="divide-y divide-slate-200/70">
        {items.map((item) => {
          const value =
            typeof item.value === "string"
              ? item.value
              : item.format === "won"
              ? won(item.value)
              : item.format === "percent"
              ? pct(item.value)
              : Number.isInteger(item.value)
              ? String(item.value)
              : item.value.toFixed(1);
          return (
            <div key={item.title} className="flex min-h-[34px] flex-wrap items-center justify-between gap-x-2 gap-y-0.5 py-1 first:pt-0 last:pb-0">
              <p className="shrink-0 break-keep text-[12px] font-semibold text-slate-500">{item.title}</p>
              <p className={`min-w-0 max-w-full flex-1 whitespace-normal break-all text-right text-[15px] font-bold leading-tight tracking-tight md:text-[17px] xl:text-[16px] 2xl:text-[17px] ${item.color || "text-slate-900"}`}>{value}</p>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function metricsByStoreType(stores: Store[], targets: TargetRecord[], ests: EstRecord[], month: string) {
  const storeCodes = new Set(stores.filter((s) => s.storeType === "매장").map((s) => s.code));

  const targetTypeOf = (t: TargetRecord) => {
    if (t.storeType === "매장" || t.storeType === "비매장") return t.storeType;
    if (t.storeCode && storeCodes.has(t.storeCode)) return "매장";
    return "비매장";
  };

  const storeTarget = targets
    .filter((t) => t.month === month && targetTypeOf(t) === "매장")
    .reduce((a, b) => a + b.amount, 0);
  const nonStoreTarget = targets
    .filter((t) => t.month === month && targetTypeOf(t) !== "매장")
    .reduce((a, b) => a + b.amount, 0);
  const storeEst = ests.filter((e) => e.month === month && storeCodes.has(e.storeCode)).reduce((a, b) => a + b.amount, 0);
  const nonStoreEst = ests.filter((e) => e.month === month && !storeCodes.has(e.storeCode)).reduce((a, b) => a + b.amount, 0);
  return { storeTarget, nonStoreTarget, storeEst, nonStoreEst };
}

function Dashboard({ stores, sales, targets, ests, month, date, timeGone, codeMappings }: { stores: Store[]; sales: SalesRecord[]; targets: TargetRecord[]; ests: EstRecord[]; month: string; date: string; timeGone: ReturnType<typeof getTimeGone>; codeMappings: StoreCodeMapping[] }) {
  const current = sales.filter((s) => s.period === "current" && inRange(s.saleDate, monthStart(month), date));
  const currentFullMonth = sales.filter((s) => s.period === "current" && inRange(s.saleDate, monthStart(month), monthEnd(month)));
  const prevMonth = sales.filter((s) => s.period === "prevMonth" && s.refMonth === month);
  const prevYear = sales.filter((s) => s.period === "prevYear" && s.refMonth === month);

  const currentSales = sum(current, "salesAmount");
  const fullMonthSales = sum(currentFullMonth, "salesAmount");
  const prevMonthSales = sum(prevMonth, "salesAmount");
  const prevYearSales = sum(prevYear, "salesAmount");
  const profitAmount = sum(current, "profitAmount");
  const profitRate = weightedProfitRate(current);
  const { storeTarget, nonStoreTarget, storeEst, nonStoreEst } = metricsByStoreType(stores, targets, ests, month);
  const targetTotal = storeTarget + nonStoreTarget;
  const estTotal = storeEst + nonStoreEst;

  const dashboardExcelRows = [
    { 항목: "기준월", 값: month },
    { 항목: "기준일", 값: date },
    { 항목: "TIME GONE", 값: pct(timeGone.timeGoneRate) },
    { 항목: "매장 Target", 값: storeTarget },
    { 항목: "비매장 Target", 값: nonStoreTarget },
    { 항목: "총 Target", 값: targetTotal },
    { 항목: "Target 달성률", 값: pct(targetTotal ? (currentSales / targetTotal) * 100 : 0) },
    { 항목: "매장 EST", 값: storeEst },
    { 항목: "비매장 EST", 값: nonStoreEst },
    { 항목: "총 EST", 값: estTotal },
    { 항목: "EST 달성률", 값: pct(estTotal ? (currentSales / estTotal) * 100 : 0) },
    { 항목: "당일까지 매출", 값: currentSales },
    { 항목: "당월 전체 매출", 값: fullMonthSales },
    { 항목: "전월 대비", 값: pct(prevMonthSales ? ((currentSales - prevMonthSales) / prevMonthSales) * 100 : 0) },
    { 항목: "전년동월 대비", 값: pct(prevYearSales ? ((currentSales - prevYearSales) / prevYearSales) * 100 : 0) },
    { 항목: "이익금액", 값: profitAmount },
    { 항목: "이익률", 값: pct(profitRate) },
  ];

  return (
    <>
      <div className="mx-auto mb-4 grid max-w-7xl grid-cols-1 gap-3 xl:grid-cols-4">
        <KpiGroup items={[
          { title: "기준월", value: month },
          { title: "기준일", value: date },
          { title: "TIME GONE", value: timeGone.timeGoneRate, format: "percent", color: "text-slate-900" },
        ]} />
        <KpiGroup items={[
          { title: "매장 Target", value: storeTarget, format: "won" },
          { title: "비매장 Target", value: nonStoreTarget, format: "won" },
          { title: "총 Target", value: targetTotal, format: "won", color: "text-slate-900" },
          { title: "Target 달성률", value: targetTotal ? (currentSales / targetTotal) * 100 : 0, format: "percent", color: "text-slate-900" },
        ]} />
        <KpiGroup items={[
          { title: "매장 EST", value: storeEst, format: "won" },
          { title: "비매장 EST", value: nonStoreEst, format: "won" },
          { title: "총 EST", value: estTotal, format: "won", color: "text-slate-900" },
          { title: "EST 달성률", value: estTotal ? (currentSales / estTotal) * 100 : 0, format: "percent", color: "text-slate-900" },
        ]} />
        <KpiGroup items={[
          { title: "당일까지 매출", value: currentSales, format: "won" },
          { title: "당월 전체 매출", value: fullMonthSales, format: "won" },
          { title: "이익금액", value: profitAmount, format: "won", color: "text-slate-900" },
          { title: "이익률", value: profitRate, format: "percent", color: "text-slate-900" },
        ]} />
      </div>

      <SalesStatus stores={stores} sales={sales} targets={targets} ests={ests} month={month} date={date} timeGone={timeGone} codeMappings={codeMappings} compact defaultView="브랜드별" />
    </>
  );
}



type ItemMetric = {
  qty: number;
  sales: number;
  cost: number;
  profit: number;
};

type ItemAnalysisRow = {
  itemCode: string;
  itemName: string;
  current: ItemMetric;
  prevMonth: ItemMetric;
  prevYear: ItemMetric;
};

function emptyItemMetric(): ItemMetric {
  return { qty: 0, sales: 0, cost: 0, profit: 0 };
}

function addItemMetric(metric: ItemMetric, row: SalesRecord) {
  metric.qty += Number(row.quantity || 0);
  metric.sales += Number(row.salesAmount || 0);
  metric.cost += Number(row.costAmount || 0);
  metric.profit += Number(row.profitAmount || 0);
}

function itemMetricDiff(now: number, base: number) {
  return now - base;
}

function itemMetricRate(now: number, base: number) {
  return base ? ((now - base) / base) * 100 : 0;
}

function itemSignedNumber(n: number, isMoney = false) {
  const rounded = Math.round(n || 0);
  const sign = rounded > 0 ? "+" : "";
  return `${sign}${isMoney ? won(rounded) : rounded.toLocaleString("ko-KR")}`;
}

function itemSignedPct(n: number) {
  const sign = n > 0 ? "+" : "";
  return Number.isFinite(n) ? `${sign}${n.toFixed(1)}%` : "-";
}

function itemMarginRate(metric: ItemMetric) {
  return metric.sales ? (metric.profit / metric.sales) * 100 : 0;
}

function itemStoreMatches(store: Store, search: string) {
  const q = search.trim().toLowerCase();
  if (!q) return true;
  return [store.brand, store.code, store.name, store.channel, store.storeType, store.manager]
    .some((v) => String(v || "").toLowerCase().includes(q));
}

function salesInRangeByStore(sales: SalesRecord[], storeCode: string, start: string, end: string) {
  return sales.filter((r) => r.storeCode === storeCode && inRange(r.saleDate, start, end));
}

function sumSalesInRangeByStore(sales: SalesRecord[], storeCode: string, start: string, end: string) {
  return salesInRangeByStore(sales, storeCode, start, end).reduce((total, r) => total + Number(r.salesAmount || 0), 0);
}

function buildItemAnalysisRows(sales: SalesRecord[], storeCode: string, currentStart: string, currentEnd: string, prevStart: string, prevEnd: string, prevYearStart: string, prevYearEnd: string): ItemAnalysisRow[] {
  const itemMap = new Map<string, ItemAnalysisRow>();
  const ensure = (code: string, name: string) => {
    const key = code || name || "미지정";
    if (!itemMap.has(key)) {
      itemMap.set(key, {
        itemCode: code || "-",
        itemName: name || "미지정",
        current: emptyItemMetric(),
        prevMonth: emptyItemMetric(),
        prevYear: emptyItemMetric(),
      });
    }
    return itemMap.get(key)!;
  };

  sales.forEach((r) => {
    if (r.storeCode !== storeCode) return;
    const row = ensure(r.itemCode, r.itemName);
    if (inRange(r.saleDate, currentStart, currentEnd)) addItemMetric(row.current, r);
    if (inRange(r.saleDate, prevStart, prevEnd)) addItemMetric(row.prevMonth, r);
    if (inRange(r.saleDate, prevYearStart, prevYearEnd)) addItemMetric(row.prevYear, r);
  });

  return Array.from(itemMap.values())
    .filter((r) => r.current.qty || r.prevMonth.qty || r.prevYear.qty || r.current.sales || r.prevMonth.sales || r.prevYear.sales)
    .sort((a, b) => b.current.sales - a.current.sales);
}

function PopupTh({ children, right = false }: { children: React.ReactNode; right?: boolean }) {
  return (
    <th className={`sticky top-0 z-50 border border-slate-300 bg-white px-3 py-2 font-bold shadow-[0_2px_0_0_#e2e8f0] ${right ? "text-right" : "text-left"}`}>
      {children}
    </th>
  );
}

function ItemAnalysis({ stores, sales, month, date }: { stores: Store[]; sales: SalesRecord[]; month: string; date: string }) {
  const [mode, setMode] = useState<"브랜드별" | "매장별">("브랜드별");
  const [searchDraft, setSearchDraft] = useState("");
  const [search, setSearch] = useState("");
  const [selectedBrand, setSelectedBrand] = useState("");
  const [selectedStoreCode, setSelectedStoreCode] = useState("");
  const [selectedItemCode, setSelectedItemCode] = useState("");
  const [itemSortMode, setItemSortMode] = useState<"품목코드 오름차순" | "전월 수량 증감률 높은순" | "전월 수량 증감률 낮은순" | "전년 수량 증감률 높은순" | "전년 수량 증감률 낮은순">("품목코드 오름차순");
  const [analysisStart, setAnalysisStart] = useState(monthStart(month));
  const [analysisEnd, setAnalysisEnd] = useState(date.startsWith(month) ? date : monthEnd(month));

  useEffect(() => {
    setAnalysisStart(monthStart(month));
    setAnalysisEnd(date.startsWith(month) ? date : monthEnd(month));
    setSelectedBrand("");
    setSelectedStoreCode("");
    setSelectedItemCode("");
  }, [month, date]);

  useEffect(() => {
    setSelectedItemCode("");
  }, [analysisStart, analysisEnd]);

  const currentStart = analysisStart <= analysisEnd ? analysisStart : analysisEnd;
  const currentEnd = analysisStart <= analysisEnd ? analysisEnd : analysisStart;
  const prevStart = sameDayPrevMonth(currentStart);
  const prevEnd = sameDayPrevMonth(currentEnd);
  const prevYearStart = sameDayPrevYear(currentStart);
  const prevYearEnd = sameDayPrevYear(currentEnd);
  const normalizedSearch = search.trim().toLowerCase();
  const storeByCode = useMemo(() => new Map(stores.map((s) => [s.code, s])), [stores]);
  const activeStores = useMemo(() => stores.filter((s) => s.status !== "거래종료" && itemStoreMatches(s, normalizedSearch)), [stores, normalizedSearch]);
  const selectedStore = storeByCode.get(selectedStoreCode);

  const brandRows = useMemo(() => {
    const map = new Map<string, { brand: string; stores: Store[]; current: number; prevMonth: number; prevYear: number }>();
    activeStores.forEach((s) => {
      const brand = s.brand || "미지정";
      if (!map.has(brand)) map.set(brand, { brand, stores: [], current: 0, prevMonth: 0, prevYear: 0 });
      const row = map.get(brand)!;
      row.stores.push(s);
      row.current += sumSalesInRangeByStore(sales, s.code, currentStart, currentEnd);
      row.prevMonth += sumSalesInRangeByStore(sales, s.code, prevStart, prevEnd);
      row.prevYear += sumSalesInRangeByStore(sales, s.code, prevYearStart, prevYearEnd);
    });
    return Array.from(map.values()).sort((a, b) => b.current - a.current);
  }, [activeStores, sales, currentStart, currentEnd, prevStart, prevEnd, prevYearStart, prevYearEnd]);

  const storeRows = useMemo(() => {
    const base = mode === "브랜드별" && selectedBrand ? activeStores.filter((s) => (s.brand || "미지정") === selectedBrand) : activeStores;
    return base.map((s) => ({
      store: s,
      current: sumSalesInRangeByStore(sales, s.code, currentStart, currentEnd),
      prevMonth: sumSalesInRangeByStore(sales, s.code, prevStart, prevEnd),
      prevYear: sumSalesInRangeByStore(sales, s.code, prevYearStart, prevYearEnd),
    })).sort((a, b) => b.current - a.current);
  }, [activeStores, selectedBrand, mode, sales, currentStart, currentEnd, prevStart, prevEnd, prevYearStart, prevYearEnd]);

  const itemRows = useMemo(() => selectedStoreCode ? buildItemAnalysisRows(sales, selectedStoreCode, currentStart, currentEnd, prevStart, prevEnd, prevYearStart, prevYearEnd) : [], [sales, selectedStoreCode, currentStart, currentEnd, prevStart, prevEnd, prevYearStart, prevYearEnd]);
  const sortedItemRows = useMemo(() => {
    const rows = [...itemRows];
    const prevRate = (r: ItemAnalysisRow) => itemMetricRate(r.current.qty, r.prevMonth.qty);
    const prevYearRate = (r: ItemAnalysisRow) => itemMetricRate(r.current.qty, r.prevYear.qty);
    if (itemSortMode === "전월 수량 증감률 높은순") return rows.sort((a, b) => prevRate(b) - prevRate(a));
    if (itemSortMode === "전월 수량 증감률 낮은순") return rows.sort((a, b) => prevRate(a) - prevRate(b));
    if (itemSortMode === "전년 수량 증감률 높은순") return rows.sort((a, b) => prevYearRate(b) - prevYearRate(a));
    if (itemSortMode === "전년 수량 증감률 낮은순") return rows.sort((a, b) => prevYearRate(a) - prevYearRate(b));
    return rows.sort((a, b) => String(a.itemCode).localeCompare(String(b.itemCode), "ko-KR", { numeric: true }));
  }, [itemRows, itemSortMode]);
  const selectedItem = itemRows.find((r) => r.itemCode === selectedItemCode);
  const detailRows = useMemo(() => {
    if (!selectedStoreCode || !selectedItemCode) return [];
    return sales.filter((r) => r.storeCode === selectedStoreCode && r.itemCode === selectedItemCode && inRange(r.saleDate, currentStart, currentEnd)).sort((a, b) => a.saleDate.localeCompare(b.saleDate));
  }, [sales, selectedStoreCode, selectedItemCode, currentStart, currentEnd]);

  function applySearch() {
    setSearch(searchDraft.trim());
    setSelectedBrand("");
    setSelectedStoreCode("");
    setSelectedItemCode("");
  }

  function backToTop() {
    setSelectedBrand("");
    setSelectedStoreCode("");
    setSelectedItemCode("");
  }

  function backToStores() {
    setSelectedStoreCode("");
    setSelectedItemCode("");
  }

  function backToItems() {
    setSelectedItemCode("");
  }

  return (
    <div className="space-y-4">
      <div className="rounded-2xl border border-slate-300 bg-white p-4 shadow-sm">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div>
            <div className="text-base font-bold text-slate-900">품목분석</div>
            <div className="mt-1 text-xs text-slate-500">현재 {currentStart} ~ {currentEnd} / 전월 {prevStart} ~ {prevEnd} / 전년 {prevYearStart} ~ {prevYearEnd}</div>
          </div>
          <div className="flex flex-wrap items-center gap-2 text-xs">
            <input
              value={searchDraft}
              onChange={(e) => { const next = e.target.value; setSearchDraft(next); if (!next.trim()) setSearch(""); }}
              onKeyDown={(e) => { if (e.key === "Enter") applySearch(); }}
              placeholder="브랜드/거래처/품목/담당자 검색 후 Enter"
              className="h-8 w-[260px] rounded-lg border border-slate-300 bg-white px-3 text-xs outline-none focus:border-blue-500"
            />
            <button type="button" onClick={applySearch} className="h-8 rounded-lg bg-slate-800 px-3 text-xs font-semibold text-white hover:bg-slate-700">검색</button>
            <label className="flex items-center gap-1 text-slate-600">시작일
              <input type="date" value={analysisStart} onChange={(e) => setAnalysisStart(e.target.value)} className="h-8 rounded-lg border border-slate-300 px-2 text-xs outline-none focus:border-blue-500" />
            </label>
            <label className="flex items-center gap-1 text-slate-600">종료일
              <input type="date" value={analysisEnd} onChange={(e) => setAnalysisEnd(e.target.value)} className="h-8 rounded-lg border border-slate-300 px-2 text-xs outline-none focus:border-blue-500" />
            </label>
          </div>
          <div className="flex rounded-xl bg-slate-100 p-1 text-sm font-semibold">
            {(["브랜드별", "매장별"] as const).map((m) => (
              <button key={m} onClick={() => { setMode(m); backToTop(); }} className={`rounded-lg px-4 py-2 ${mode === m ? "bg-white text-blue-700 shadow-sm" : "text-slate-600 hover:text-blue-700"}`}>{m}</button>
            ))}
          </div>
        </div>
      </div>

      {!selectedStoreCode && mode === "브랜드별" && !selectedBrand && (
        <div className="overflow-hidden rounded-2xl border border-slate-300 bg-white shadow-sm">
          <div className="border-b border-slate-300 px-4 py-3 text-sm font-bold text-slate-800">브랜드별 요약</div>
          <div className="max-h-[65vh] overflow-auto isolate">
            <table className="w-full min-w-[900px] border-separate border-spacing-0 text-center text-sm">
              <thead><tr>{["브랜드", "당일까지 매출", "전월 매출", "전월 대비", "전년 매출", "전년 대비", "거래처수", "상세"].map((h) => <PopupTh key={h}>{h}</PopupTh>)}</tr></thead>
              <tbody>
                {brandRows.map((r) => (
                  <tr key={r.brand} className="hover:bg-blue-50">
                    <td className="border border-slate-300 p-2 font-semibold">{r.brand}</td>
                    <td className="border border-slate-300 p-2 text-right font-bold text-blue-700">{won(r.current)}</td>
                    <td className="border border-slate-300 p-2 text-right">{won(r.prevMonth)}</td>
                    <td className={`border border-slate-300 p-2 text-right ${itemMetricDiff(r.current, r.prevMonth) >= 0 ? "text-emerald-600" : "text-red-600"}`}>{itemSignedPct(itemMetricRate(r.current, r.prevMonth))}</td>
                    <td className="border border-slate-300 p-2 text-right">{won(r.prevYear)}</td>
                    <td className={`border border-slate-300 p-2 text-right ${itemMetricDiff(r.current, r.prevYear) >= 0 ? "text-emerald-600" : "text-red-600"}`}>{itemSignedPct(itemMetricRate(r.current, r.prevYear))}</td>
                    <td className="border border-slate-300 p-2 text-right">{r.stores.length}</td>
                    <td className="border border-slate-300 p-2"><button onClick={() => setSelectedBrand(r.brand)} className="rounded-lg bg-blue-600 px-3 py-1 text-xs font-semibold text-white">거래처 보기</button></td>
                  </tr>
                ))}
                {!brandRows.length && <tr><td colSpan={8} className="border border-slate-300 p-8 text-center text-slate-500">표시할 데이터가 없습니다.</td></tr>}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {!selectedStoreCode && (mode === "매장별" || selectedBrand) && (
        <div className="overflow-hidden rounded-2xl border border-slate-300 bg-white shadow-sm">
          <div className="flex items-center justify-between border-b border-slate-300 px-4 py-3">
            <div className="text-sm font-bold text-slate-800">{selectedBrand ? `${selectedBrand} 거래처 목록` : "거래처별 요약"}</div>
            {selectedBrand && <button onClick={backToTop} className="rounded-lg bg-slate-100 px-3 py-1 text-xs font-semibold text-slate-700 hover:bg-slate-200">← 브랜드 목록</button>}
          </div>
          <div className="max-h-[65vh] overflow-auto isolate">
            <table className="w-full min-w-[1000px] border-separate border-spacing-0 text-center text-sm">
              <thead><tr>{["브랜드", "거래처코드", "거래처명", "담당자", "당일까지 매출", "전월 대비", "전년 대비", "상세"].map((h) => <PopupTh key={h}>{h}</PopupTh>)}</tr></thead>
              <tbody>
                {storeRows.map((r) => (
                  <tr key={r.store.code} className="hover:bg-blue-50">
                    <td className="border border-slate-300 p-2">{r.store.brand}</td>
                    <td className="border border-slate-300 p-2">{r.store.code}</td>
                    <td className="border border-slate-300 p-2 font-semibold">{r.store.name}</td>
                    <td className="border border-slate-300 p-2">{r.store.manager || "미지정"}</td>
                    <td className="border border-slate-300 p-2 text-right font-bold text-blue-700">{won(r.current)}</td>
                    <td className={`border border-slate-300 p-2 text-right ${itemMetricDiff(r.current, r.prevMonth) >= 0 ? "text-emerald-600" : "text-red-600"}`}>{itemSignedPct(itemMetricRate(r.current, r.prevMonth))}</td>
                    <td className={`border border-slate-300 p-2 text-right ${itemMetricDiff(r.current, r.prevYear) >= 0 ? "text-emerald-600" : "text-red-600"}`}>{itemSignedPct(itemMetricRate(r.current, r.prevYear))}</td>
                    <td className="border border-slate-300 p-2"><button onClick={() => setSelectedStoreCode(r.store.code)} className="rounded-lg bg-blue-600 px-3 py-1 text-xs font-semibold text-white">품목 보기</button></td>
                  </tr>
                ))}
                {!storeRows.length && <tr><td colSpan={8} className="border border-slate-300 p-8 text-center text-slate-500">표시할 데이터가 없습니다.</td></tr>}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {selectedStoreCode && !selectedItemCode && selectedStore && (
        <div className="overflow-hidden rounded-2xl border border-slate-300 bg-white shadow-sm">
          <div className="flex items-center justify-between border-b border-slate-300 px-4 py-3">
            <div>
              <div className="text-sm font-bold text-slate-800">{selectedStore.name} 품목별 비교</div>
              <div className="mt-1 text-xs text-slate-500">현재/전월/전년 동일기간 기준으로 수량과 매출을 비교합니다.</div>
            </div>
            <div className="flex items-center gap-2">
              <select value={itemSortMode} onChange={(e) => setItemSortMode(e.target.value as typeof itemSortMode)} className="h-8 rounded-lg border border-slate-300 bg-white px-2 text-xs outline-none focus:border-blue-500">
                <option value="품목코드 오름차순">품목코드 오름차순</option>
                <option value="전월 수량 증감률 높은순">전월 수량 증감률 높은순</option>
                <option value="전월 수량 증감률 낮은순">전월 수량 증감률 낮은순</option>
                <option value="전년 수량 증감률 높은순">전년 수량 증감률 높은순</option>
                <option value="전년 수량 증감률 낮은순">전년 수량 증감률 낮은순</option>
              </select>
              <button onClick={backToStores} className="rounded-lg bg-slate-100 px-3 py-1 text-xs font-semibold text-slate-700 hover:bg-slate-200">← 거래처 목록</button>
            </div>
          </div>
          <div className="max-h-[65vh] overflow-auto isolate">
            <table className="w-full min-w-[1500px] border-separate border-spacing-0 text-center text-xs">
              <thead><tr>{["품목코드", "품목명", "현재수량", "전월수량", "수량차이", "수량증감%", "전년수량", "수량차이", "수량증감%", "현재매출", "전월매출", "매출차이", "매출증감%", "전년매출", "매출차이", "매출증감%", "상세"].map((h) => <PopupTh key={h}>{h}</PopupTh>)}</tr></thead>
              <tbody>
                {sortedItemRows.map((r) => (
                  <tr key={`${r.itemCode}-${r.itemName}`} className="hover:bg-blue-50">
                    <td className="border border-slate-300 p-2">{r.itemCode}</td>
                    <td className="border border-slate-300 p-2 font-semibold">{r.itemName}</td>
                    <td className="border border-slate-300 p-2 text-right font-bold text-blue-700">{won(r.current.qty)}</td>
                    <td className="border border-slate-300 p-2 text-right">{won(r.prevMonth.qty)}</td>
                    <td className={`border border-slate-300 p-2 text-right ${itemMetricDiff(r.current.qty, r.prevMonth.qty) >= 0 ? "text-emerald-600" : "text-red-600"}`}>{itemSignedNumber(itemMetricDiff(r.current.qty, r.prevMonth.qty))}</td>
                    <td className={`border border-slate-300 p-2 text-right ${itemMetricDiff(r.current.qty, r.prevMonth.qty) >= 0 ? "text-emerald-600" : "text-red-600"}`}>{itemSignedPct(itemMetricRate(r.current.qty, r.prevMonth.qty))}</td>
                    <td className="border border-slate-300 p-2 text-right">{won(r.prevYear.qty)}</td>
                    <td className={`border border-slate-300 p-2 text-right ${itemMetricDiff(r.current.qty, r.prevYear.qty) >= 0 ? "text-emerald-600" : "text-red-600"}`}>{itemSignedNumber(itemMetricDiff(r.current.qty, r.prevYear.qty))}</td>
                    <td className={`border border-slate-300 p-2 text-right ${itemMetricDiff(r.current.qty, r.prevYear.qty) >= 0 ? "text-emerald-600" : "text-red-600"}`}>{itemSignedPct(itemMetricRate(r.current.qty, r.prevYear.qty))}</td>
                    <td className="border border-slate-300 p-2 text-right font-bold text-blue-700">{won(r.current.sales)}</td>
                    <td className="border border-slate-300 p-2 text-right">{won(r.prevMonth.sales)}</td>
                    <td className={`border border-slate-300 p-2 text-right ${itemMetricDiff(r.current.sales, r.prevMonth.sales) >= 0 ? "text-emerald-600" : "text-red-600"}`}>{itemSignedNumber(itemMetricDiff(r.current.sales, r.prevMonth.sales), true)}</td>
                    <td className={`border border-slate-300 p-2 text-right ${itemMetricDiff(r.current.sales, r.prevMonth.sales) >= 0 ? "text-emerald-600" : "text-red-600"}`}>{itemSignedPct(itemMetricRate(r.current.sales, r.prevMonth.sales))}</td>
                    <td className="border border-slate-300 p-2 text-right">{won(r.prevYear.sales)}</td>
                    <td className={`border border-slate-300 p-2 text-right ${itemMetricDiff(r.current.sales, r.prevYear.sales) >= 0 ? "text-emerald-600" : "text-red-600"}`}>{itemSignedNumber(itemMetricDiff(r.current.sales, r.prevYear.sales), true)}</td>
                    <td className={`border border-slate-300 p-2 text-right ${itemMetricDiff(r.current.sales, r.prevYear.sales) >= 0 ? "text-emerald-600" : "text-red-600"}`}>{itemSignedPct(itemMetricRate(r.current.sales, r.prevYear.sales))}</td>
                    <td className="border border-slate-300 p-2"><button onClick={() => setSelectedItemCode(r.itemCode)} className="rounded-lg bg-slate-800 px-3 py-1 text-xs font-semibold text-white">상세</button></td>
                  </tr>
                ))}
                {!sortedItemRows.length && <tr><td colSpan={17} className="border border-slate-300 p-8 text-center text-slate-500">표시할 품목이 없습니다.</td></tr>}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {selectedStoreCode && selectedItemCode && selectedItem && selectedStore && (
        <div className="space-y-4">
          <div className="rounded-2xl border border-slate-300 bg-white p-4 shadow-sm">
            <div className="flex items-center justify-between gap-3">
              <div>
                <div className="text-sm font-bold text-slate-800">{selectedStore.name} / {selectedItem.itemName}</div>
                <div className="mt-1 text-xs text-slate-500">품목코드 {selectedItem.itemCode} / 현재기간 상세 발주 원본</div>
              </div>
              <button onClick={backToItems} className="rounded-lg bg-slate-100 px-3 py-1 text-xs font-semibold text-slate-700 hover:bg-slate-200">← 품목 목록</button>
            </div>
          </div>
          <div className="grid gap-4 lg:grid-cols-3">
            <ItemCompareCard title="수량" current={selectedItem.current.qty} prev={selectedItem.prevMonth.qty} prevYear={selectedItem.prevYear.qty} />
            <ItemCompareCard title="매출" current={selectedItem.current.sales} prev={selectedItem.prevMonth.sales} prevYear={selectedItem.prevYear.sales} money />
            <div className="rounded-2xl border border-slate-300 bg-white p-4 text-sm shadow-sm">
              <div className="font-bold text-slate-800">이익</div>
              <div className="mt-3 space-y-2 text-xs">
                <div className="flex justify-between"><span>현재 이익금액</span><b>{won(selectedItem.current.profit)}</b></div>
                <div className="flex justify-between"><span>현재 이익률</span><b>{pct(itemMarginRate(selectedItem.current))}</b></div>
                <div className="flex justify-between"><span>전월 이익금액</span><b>{won(selectedItem.prevMonth.profit)}</b></div>
                <div className="flex justify-between"><span>전월 이익률</span><b>{pct(itemMarginRate(selectedItem.prevMonth))}</b></div>
                <div className="flex justify-between"><span>전년 이익금액</span><b>{won(selectedItem.prevYear.profit)}</b></div>
                <div className="flex justify-between"><span>전년 이익률</span><b>{pct(itemMarginRate(selectedItem.prevYear))}</b></div>
              </div>
            </div>
          </div>
          <div className="overflow-hidden rounded-2xl border border-slate-300 bg-white shadow-sm">
            <div className="border-b border-slate-300 px-4 py-3 text-sm font-bold text-slate-800">상세 발주 원본</div>
            <div className="max-h-[55vh] overflow-auto isolate">
              <table className="w-full min-w-[1100px] border-separate border-spacing-0 text-center text-xs">
                <thead><tr>{["주문일", "거래처", "상품코드", "상품명", "수량", "매출금액", "원가금액", "이익금액", "이익률"].map((h, idx) => <PopupTh key={h} right={idx >= 4}>{h}</PopupTh>)}</tr></thead>
                <tbody>
                  {detailRows.map((r) => (
                    <tr key={r.id} className="hover:bg-blue-50">
                      <td className="border border-slate-300 p-2">{r.saleDate}</td>
                      <td className="border border-slate-300 p-2 font-semibold">{r.storeName}</td>
                      <td className="border border-slate-300 p-2">{r.itemCode}</td>
                      <td className="border border-slate-300 p-2">{r.itemName}</td>
                      <td className="border border-slate-300 p-2 text-right">{won(r.quantity)}</td>
                      <td className="border border-slate-300 p-2 text-right font-bold text-slate-900">{won(r.salesAmount)}</td>
                      <td className="border border-slate-300 p-2 text-right">{won(r.costAmount)}</td>
                      <td className="border border-slate-300 p-2 text-right">{won(r.profitAmount)}</td>
                      <td className="border border-slate-300 p-2 text-right">{pct(r.profitRate)}</td>
                    </tr>
                  ))}
                  {!detailRows.length && <tr><td colSpan={9} className="border border-slate-300 p-8 text-center text-slate-500">상세 데이터가 없습니다.</td></tr>}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function ItemCompareCard({ title, current, prev, prevYear, money }: { title: string; current: number; prev: number; prevYear: number; money?: boolean }) {
  return (
    <div className="rounded-2xl border border-slate-300 bg-white p-4 text-sm shadow-sm">
      <div className="font-bold text-slate-800">{title}</div>
      <div className="mt-3 space-y-2 text-xs">
        <div className="flex justify-between"><span>현재</span><b className="text-blue-700">{money ? won(current) : won(current)}</b></div>
        <div className="flex justify-between"><span>전월</span><b>{money ? won(prev) : won(prev)}</b></div>
        <div className="flex justify-between"><span>전월 차이</span><b className={itemMetricDiff(current, prev) >= 0 ? "text-emerald-600" : "text-red-600"}>{itemSignedNumber(itemMetricDiff(current, prev), money)}</b></div>
        <div className="flex justify-between"><span>전월 증감</span><b className={itemMetricDiff(current, prev) >= 0 ? "text-emerald-600" : "text-red-600"}>{itemSignedPct(itemMetricRate(current, prev))}</b></div>
        <div className="flex justify-between"><span>전년</span><b>{money ? won(prevYear) : won(prevYear)}</b></div>
        <div className="flex justify-between"><span>전년 차이</span><b className={itemMetricDiff(current, prevYear) >= 0 ? "text-emerald-600" : "text-red-600"}>{itemSignedNumber(itemMetricDiff(current, prevYear), money)}</b></div>
        <div className="flex justify-between"><span>전년 증감</span><b className={itemMetricDiff(current, prevYear) >= 0 ? "text-emerald-600" : "text-red-600"}>{itemSignedPct(itemMetricRate(current, prevYear))}</b></div>
      </div>
    </div>
  );
}

function SalesStatus({ stores, sales, targets, ests, month, date, timeGone, codeMappings, compact = false, defaultView = "거래처별" }: { stores: Store[]; sales: SalesRecord[]; targets: TargetRecord[]; ests: EstRecord[]; month: string; date: string; timeGone: ReturnType<typeof getTimeGone>; codeMappings: StoreCodeMapping[]; compact?: boolean; defaultView?: SalesView }) {
  const [view, setView] = useState<SalesView>(defaultView);
  const [search, setSearch] = useState("");
  const [searchDraft, setSearchDraft] = useState("");
  const [drill, setDrill] = useState<{ title: string; rows: SalesRecord[] } | null>(null);
  const [sortConfig, setSortConfig] = useState<{ key: SalesStatusSortKey; direction: SortDirection }>({ key: "currentSales", direction: "desc" });
  const [inactiveOpen, setInactiveOpen] = useState(false);
  const [orderDateFilter, setOrderDateFilter] = useState<"all" | "check7" | "no30">("all");
  const [hideEndedStores, setHideEndedStores] = useState(false);

  const normalizedSearch = search.trim().toLowerCase();
  const stMap = storeMap(stores);

  const currentCanonicalStores = useMemo(() => {
    const map = new Map<string, ReturnType<typeof resolveStoreInfo>>();
    sales
      .filter((r) => r.period === "current" && inRange(r.saleDate, monthStart(month), monthEnd(month)))
      .forEach((r) => {
        if (!map.has(r.storeCode)) {
          map.set(r.storeCode, resolveStoreInfo(r.storeCode, r.storeName, {
            channel: r.channel,
            manager: r.manager,
            storeType: r.storeType,
            brand: r.brand,
          }, stores));
        }
      });
    return Array.from(map.values());
  }, [sales, month, stores]);

  const currentByCode = useMemo(() => new Map(currentCanonicalStores.map((s) => [norm(s.code), s])), [currentCanonicalStores]);
  const currentByName = useMemo(() => {
    const map = new Map<string, ReturnType<typeof resolveStoreInfo>>();
    currentCanonicalStores.forEach((s) => {
      const key = normalizeStoreNameKey(s.name);
      if (key && !map.has(key)) map.set(key, s);
    });
    return map;
  }, [currentCanonicalStores]);

  const findManualMapping = (code: string, name: string) => {
    const codeKey = norm(code);
    const nameKey = normalizeStoreNameKey(name);
    return codeMappings.find((m) => {
      if (norm(m.oldCode) !== codeKey) return false;
      const mappedOldNameKey = normalizeStoreNameKey(m.oldName);
      return !mappedOldNameKey || mappedOldNameKey === nameKey;
    });
  };

  const uploadedStoreInfo = (r: SalesRecord) => {
    const mappedStore = stMap.get(r.storeCode);
    return {
      code: mappedStore?.code || r.storeCode || r.storeName || "미지정",
      name: mappedStore?.name || r.storeName || r.storeCode || "미지정",
      channel: mappedStore?.channel || r.channel || "미지정",
      manager: mappedStore?.manager || r.manager || "미지정",
      storeType: mappedStore?.storeType || r.storeType || "비매장",
      brand: mappedStore?.brand || r.brand || "미지정",
      status: mappedStore?.status || ("거래중" as const),
      originalCode: r.storeCode,
      originalName: r.storeName,
    };
  };

  const resolveRecord = (r: SalesRecord) => {
    if (r.period === "prevYear" || r.period === "prevMonth") {
      const manual = findManualMapping(r.storeCode, r.storeName);
      if (manual?.currentCode) {
        const mapped = resolveStoreInfo(manual.currentCode, manual.currentName || r.storeName, {
          channel: r.channel,
          manager: r.manager,
          storeType: r.storeType,
          brand: r.brand,
        }, stores);
        return {
          ...mapped,
          originalCode: r.storeCode,
          originalName: r.storeName,
        };
      }

      const byCode = currentByCode.get(norm(r.storeCode));
      const byName = currentByName.get(normalizeStoreNameKey(r.storeName));
      const autoTarget = byCode || byName;
      if (autoTarget) {
        return {
          ...autoTarget,
          originalCode: r.storeCode,
          originalName: r.storeName,
        };
      }

      return uploadedStoreInfo(r);
    }

    return resolveStoreInfo(r.storeCode, r.storeName, {
      channel: r.channel,
      manager: r.manager,
      storeType: r.storeType,
      brand: r.brand,
    }, stores);
  };

  const filterByStoreSearch = (s: SalesRecord) => {
    if (!normalizedSearch) return true;

    const resolved = resolveRecord(s);

    // 거래처별 화면에서는 검색어가 거래처코드/거래처명에 걸린 행만 보여줍니다.
    // 브랜드, 담당자, 채널까지 검색에 포함하면 검색어와 직접 관련 없는 거래처 행까지 살아남을 수 있습니다.
    if (view === "거래처별") {
      return (
        s.storeName.toLowerCase().includes(normalizedSearch) ||
        s.storeCode.toLowerCase().includes(normalizedSearch) ||
        resolved.name.toLowerCase().includes(normalizedSearch) ||
        resolved.code.toLowerCase().includes(normalizedSearch)
      );
    }

    return (
      s.storeName.toLowerCase().includes(normalizedSearch) ||
      s.storeCode.toLowerCase().includes(normalizedSearch) ||
      s.brand.toLowerCase().includes(normalizedSearch) ||
      String(s.manager || "").toLowerCase().includes(normalizedSearch) ||
      s.channel.toLowerCase().includes(normalizedSearch) ||
      resolved.name.toLowerCase().includes(normalizedSearch) ||
      resolved.code.toLowerCase().includes(normalizedSearch) ||
      resolved.brand.toLowerCase().includes(normalizedSearch) ||
      String(resolved.manager || "").toLowerCase().includes(normalizedSearch) ||
      resolved.channel.toLowerCase().includes(normalizedSearch)
    );
  };

  const shouldIncludeRecord = (s: SalesRecord) => !hideEndedStores || resolveRecord(s).status !== "거래종료";

  const current = sales.filter((s) => s.period === "current" && inRange(s.saleDate, monthStart(month), date) && filterByStoreSearch(s) && shouldIncludeRecord(s));
  const currentFullMonthRows = sales.filter((s) => s.period === "current" && inRange(s.saleDate, monthStart(month), monthEnd(month)) && filterByStoreSearch(s) && shouldIncludeRecord(s));
  const prevMonthRows = sales.filter((s) => s.period === "prevMonth" && s.refMonth === month && filterByStoreSearch(s) && shouldIncludeRecord(s));
  const prevYearRows = sales.filter((s) => s.period === "prevYear" && s.refMonth === month && filterByStoreSearch(s) && shouldIncludeRecord(s));

  const rowKey = (r: SalesRecord) => {
    const resolved = resolveRecord(r);
    if (view === "거래처별") return resolved.code || resolved.name;
    if (view === "브랜드별") return resolved.brand || "미지정";
    if (view === "담당자별") return resolved.manager || "미지정";
    return resolved.channel || "미지정";
  };

  const rowLabel = (key: string, records: SalesRecord[]) => {
    if (view !== "거래처별") return key || "미지정";
    const first = records[0];
    if (first) return resolveRecord(first).name;
    const mapped = stMap.get(key);
    return mapped?.name || key || "미지정";
  };

  const storeKey = (store: Store) => {
    const resolved = resolveStoreInfo(store.code, store.name, store, stores);
    if (view === "거래처별") return resolved.code || resolved.name;
    if (view === "브랜드별") return resolved.brand || "미지정";
    if (view === "담당자별") return resolved.manager || "미지정";
    return resolved.channel || "미지정";
  };

  const currentMap = groupByKey(current, rowKey);
  const currentFullMonthMap = groupByKey(currentFullMonthRows, rowKey);
  const prevMonthMap = groupByKey(prevMonthRows, rowKey);
  const prevYearMap = groupByKey(prevYearRows, rowKey);

  const estMap = new Map<string, number>();
  ests
    .filter((e) => {
      const mappedStore = stMap.get(e.storeCode);
      const resolved = resolveStoreInfo(e.storeCode, e.storeName, mappedStore || {}, stores);
      const display = resolved;
      const mappedStoreName = `${mappedStore?.name || ""} ${display.name}`;
      const mappedBrand = `${mappedStore?.brand || ""} ${display.brand}`;
      const mappedManager = `${mappedStore?.manager || ""} ${display.manager}`;
      const mappedChannel = `${mappedStore?.channel || ""} ${display.channel}`;
      const isEnded = display.status === "거래종료" || mappedStore?.status === "거래종료";
      if (e.month !== month || (hideEndedStores && isEnded)) return false;
      if (!normalizedSearch) return true;

      if (view === "거래처별") {
        return (
          e.storeName.toLowerCase().includes(normalizedSearch) ||
          e.storeCode.toLowerCase().includes(normalizedSearch) ||
          mappedStoreName.toLowerCase().includes(normalizedSearch)
        );
      }

      return (
        e.storeName.toLowerCase().includes(normalizedSearch) ||
        e.storeCode.toLowerCase().includes(normalizedSearch) ||
        mappedStoreName.toLowerCase().includes(normalizedSearch) ||
        mappedBrand.toLowerCase().includes(normalizedSearch) ||
        mappedManager.toLowerCase().includes(normalizedSearch) ||
        mappedChannel.toLowerCase().includes(normalizedSearch)
      );
    })
    .forEach((e) => {
      const s = stMap.get(e.storeCode);
      const resolved = resolveStoreInfo(e.storeCode, e.storeName, s || {}, stores);
      const display = resolved;
      const key = s ? storeKey(s) : view === "거래처별" ? display.code || display.name || "미지정" : view === "브랜드별" ? display.brand || "미지정" : view === "담당자별" ? display.manager || "미지정" : display.channel || "미지정";
      estMap.set(key, (estMap.get(key) || 0) + e.amount);
    });

  const baseKeySet = new Set([...currentMap.keys(), ...currentFullMonthMap.keys(), ...prevMonthMap.keys(), ...prevYearMap.keys(), ...estMap.keys()]);
  if (view === "거래처별" && !normalizedSearch) {
    stores.filter((store) => !hideEndedStores || store.status === "거래중").forEach((store) => baseKeySet.add(storeKey(store)));
  }
  const keys = Array.from(baseKeySet).sort();

  const lastOrderDateByKey = useMemo(() => {
    const map = new Map<string, string>();
    sales
      .filter((r) => r.period === "current" && r.saleDate <= date && shouldIncludeRecord(r))
      .forEach((r) => {
        const key = rowKey(r);
        const prev = map.get(key);
        if (!prev || r.saleDate > prev) map.set(key, r.saleDate);
      });
    return map;
  }, [sales, date, view, stores, codeMappings]);

  const getDrillRows = (key: string, period: DrillPeriod) => {
    if (period === "prevYear") return prevYearMap.get(key) || [];
    if (period === "prevMonth") return prevMonthMap.get(key) || [];
    if (period === "current") return currentMap.get(key) || [];
    return currentFullMonthMap.get(key) || [];
  };

  const getTotalDrillRows = (period: DrillPeriod) => {
    if (period === "prevYear") return prevYearRows;
    if (period === "prevMonth") return prevMonthRows;
    if (period === "current") return current;
    return currentFullMonthRows;
  };

  const rows = keys.map((key) => {
    const currentRecords = currentMap.get(key) || [];
    const fullMonthRecords = currentFullMonthMap.get(key) || [];
    const prevMonthRecords = prevMonthMap.get(key) || [];
    const prevYearRecords = prevYearMap.get(key) || [];
    const allRecords = [...currentRecords, ...fullMonthRecords, ...prevMonthRecords, ...prevYearRecords];
    const currentSales = sum(currentRecords, "salesAmount");
    const fullMonthSales = sum(fullMonthRecords, "salesAmount");
    const prevMonthSales = sum(prevMonthRecords, "salesAmount");
    const prevYearSales = sum(prevYearRecords, "salesAmount");
    const profitAmount = sum(currentRecords, "profitAmount");
    const profitRate = weightedProfitRate(currentRecords);
    const prevMonthRate = prevMonthSales ? ((currentSales - prevMonthSales) / prevMonthSales) * 100 : 0;
    const prevYearRate = prevYearSales ? ((currentSales - prevYearSales) / prevYearSales) * 100 : 0;
    const est = estMap.get(key) || 0;
    const estRate = est ? (currentSales / est) * 100 : 0;
    const firstRecord = allRecords[0];
    const storeStatus = firstRecord ? resolveRecord(firstRecord).status : stMap.get(key)?.status;
    const isEndedStore = view === "거래처별" && storeStatus === "거래종료";
    return {
      key,
      label: rowLabel(key, allRecords),
      prevYearSales,
      prevMonthSales,
      currentSales,
      fullMonthSales,
      prevMonthRate,
      prevYearRate,
      timeGone: timeGone.timeGoneRate,
      timeGoneGap: estRate - timeGone.timeGoneRate,
      est,
      estRate,
      profitAmount,
      profitRate,
      isEndedStore,
      lastOrderDate: view === "거래처별" ? (lastOrderDateByKey.get(key) || "-") : "-",
      daysSinceLastOrder: view === "거래처별" ? daysBetween(lastOrderDateByKey.get(key) || "-", date) : 0,
    };
  }).filter((row) => {
    if (view !== "거래처별") return true;
    if (orderDateFilter === "check7") return row.daysSinceLastOrder >= 7;
    if (orderDateFilter === "no30") return row.daysSinceLastOrder >= 30;
    return true;
  });

  const displayRows = useMemo(() => {
    if (!hideEndedStores || view !== "거래처별") return rows;
    return rows.filter((row) => !row.isEndedStore);
  }, [rows, hideEndedStores, view]);

  const sortedRows = useMemo(() => {
    return [...displayRows].sort((a, b) => {
      const aValue = a[sortConfig.key];
      const bValue = b[sortConfig.key];
      let result = 0;
      if (typeof aValue === "string" || typeof bValue === "string") {
        result = String(aValue).localeCompare(String(bValue), "ko-KR");
      } else {
        result = Number(aValue || 0) - Number(bValue || 0);
      }
      return sortConfig.direction === "asc" ? result : -result;
    });
  }, [displayRows, sortConfig]);

  const requestSort = (key: SalesStatusSortKey) => {
    setSortConfig((prev) => ({
      key,
      direction: prev.key === key && prev.direction === "desc" ? "asc" : "desc",
    }));
  };

  const openDrill = (row: { key: string; label: string }, period: DrillPeriod) => {
    const rows = getDrillRows(row.key, period);
    setDrill({ title: `${row.label} · ${drillPeriodLabel(period)}`, rows });
  };

  const openTotalDrill = (period: DrillPeriod) => {
    const rows = getTotalDrillRows(period);
    setDrill({ title: `${view} 전체 · ${drillPeriodLabel(period)}`, rows });
  };

  const filteredCurrentSales = displayRows.reduce((a, b) => a + b.currentSales, 0);
  const filteredFullMonthSales = displayRows.reduce((a, b) => a + b.fullMonthSales, 0);
  const filteredPrevMonthSales = displayRows.reduce((a, b) => a + b.prevMonthSales, 0);
  const filteredPrevYearSales = displayRows.reduce((a, b) => a + b.prevYearSales, 0);
  const filteredProfitAmount = displayRows.reduce((a, b) => a + b.profitAmount, 0);

  const salesStatusExcelRows = sortedRows.map((r) => ({
    구분: r.label,
    전년동월: r.prevYearSales,
    전년대비: pct(r.prevYearRate),
    전월: r.prevMonthSales,
    전월대비: pct(r.prevMonthRate),
    "당일까지 매출": r.currentSales,
    "당월 전체 매출": r.fullMonthSales,
    "TIME GONE 대비": pct(r.timeGoneGap),
    EST: r.est,
    "EST 달성률": pct(r.estRate),
    이익금액: r.profitAmount,
    이익률: pct(r.profitRate),
  }));

  return (
    <>
      <div className="rounded-2xl border border-slate-200/70 bg-white/80 p-4 shadow-sm backdrop-blur">
        <div className="mb-3 flex flex-col gap-3 lg:flex-row lg:items-center lg:justify-between">
          <div>
            <h2 className="text-base font-bold text-slate-900">매출현황</h2>
            <p className="mt-1 text-xs text-slate-500">상단 헤더를 클릭하면 내림차순/오름차순으로 정렬됩니다. 금액을 클릭하면 해당 주문내역을 볼 수 있습니다.</p>
          </div>
          <div className="ml-auto flex flex-wrap items-center justify-end gap-2">
            {compact && (
              <select
                value={view}
                onChange={(e) => setView(e.target.value as SalesView)}
                className="w-[150px] rounded-lg border border-slate-300 bg-white px-3 py-2 text-xs font-semibold text-slate-700 shadow-sm outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
              >
                {SALES_VIEWS.map((v) => (
                  <option key={v} value={v}>{v}</option>
                ))}
              </select>
            )}
            <label className="flex items-center gap-2 rounded-lg border border-slate-300 bg-white px-3 py-2 text-xs font-semibold text-slate-700 shadow-sm">
              <input
                type="checkbox"
                checked={hideEndedStores}
                onChange={(e) => setHideEndedStores(e.target.checked)}
                className="h-4 w-4 rounded border-slate-300"
              />
              거래종료 거래처 제외
            </label>
          </div>
        </div>

        {!compact && (
          <>
            <div className="mb-3 flex flex-wrap items-center gap-3">
              <select
                value={view}
                onChange={(e) => setView(e.target.value as SalesView)}
                className="w-[220px] rounded-lg border border-slate-300 bg-white/80 px-3 py-1.5 text-xs font-semibold outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
              >
                {SALES_VIEWS.map((v) => (
                  <option key={v} value={v}>{v}</option>
                ))}
              </select>
              <input
                value={searchDraft}
                onChange={(e) => {
                  const value = e.target.value;
                  setSearchDraft(value);
                  if (!value.trim()) setSearch("");
                }}
                onKeyDown={(e) => {
                  if (e.key === "Enter") setSearch(searchDraft.trim());
                }}
                placeholder="거래처 검색"
                className="w-[220px] rounded-lg border border-slate-300 bg-white/80 px-3 py-1.5 text-xs outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
              />
              {view === "거래처별" && (
                <select
                  value={orderDateFilter}
                  onChange={(e) => setOrderDateFilter(e.target.value as "all" | "check7" | "no30")}
                  className="w-[190px] rounded-lg border border-slate-300 bg-white/80 px-3 py-1.5 text-xs font-semibold outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                >
                  <option value="all">마지막 발주일 전체</option>
                  <option value="check7">확인필요 7일 이상</option>
                  <option value="no30">미발주 30일 이상</option>
                </select>
              )}
              <span className="text-xs font-medium text-slate-500">표시 기준: {view}</span>
              <div className="ml-auto flex flex-wrap items-center justify-end gap-2">
                <button
                  type="button"
                  onClick={() => setInactiveOpen(true)}
                  className="rounded-lg border border-slate-300 bg-white px-3 py-1.5 text-xs font-semibold text-slate-700 shadow-sm hover:bg-slate-50"
                >
                  3개월 미주문 보기
                </button>
                <button
                  onClick={() => exportExcel(salesStatusExcelRows, `매출현황_${month}_${search || "전체"}_${view}`)}
                  className="rounded-lg bg-emerald-600 px-3 py-1.5 text-xs font-semibold text-white shadow-sm hover:bg-emerald-700"
                >
                  매출현황 엑셀 다운로드
                </button>
              </div>
            </div>

            <div className="mb-3 grid grid-cols-2 gap-3 lg:grid-cols-5">
              <FilterAmountCard title="전년동월 매출" value={filteredPrevYearSales} tone="mint" onClick={() => openTotalDrill("prevYear")} />
              <FilterAmountCard title="전월 매출" value={filteredPrevMonthSales} tone="blue" onClick={() => openTotalDrill("prevMonth")} />
              <FilterAmountCard title="당일까지의 매출" value={filteredCurrentSales} tone="yellow" onClick={() => openTotalDrill("current")} />
              <FilterAmountCard title="당월 발주 총 금액" value={filteredFullMonthSales} tone="yellow" onClick={() => openTotalDrill("currentFullMonth")} />
              <FilterAmountCard title="이익 금액" value={filteredProfitAmount} tone="green" />
            </div>
          </>
        )}

        <div className="relative max-h-[62vh] overflow-auto bg-white">
          <table className="w-full min-w-[1320px] table-auto border-separate border-spacing-0 border border-slate-200 text-[12px] leading-tight">
            <thead>
              <tr className="bg-slate-100">
                <ThCompactSortable w="w-[10%]" sortKey="label" sortConfig={sortConfig} onSort={requestSort}>{view.replace("별", "")}</ThCompactSortable>
                {!compact && view === "거래처별" && <ThCompact tone="gray">마지막 발주일</ThCompact>}
                <ThCompactSortable right tone="mint" sortKey="prevYearSales" sortConfig={sortConfig} onSort={requestSort}>전년동월</ThCompactSortable>
                <ThCompactSortable right tone="mint" sortKey="prevYearRate" sortConfig={sortConfig} onSort={requestSort}>전년대비</ThCompactSortable>
                <ThCompactSortable right tone="blue" sortKey="prevMonthSales" sortConfig={sortConfig} onSort={requestSort}>전월</ThCompactSortable>
                <ThCompactSortable right tone="blue" sortKey="prevMonthRate" sortConfig={sortConfig} onSort={requestSort}>전월대비</ThCompactSortable>
                <ThCompactSortable right tone="yellow" sortKey="currentSales" sortConfig={sortConfig} onSort={requestSort}>당일까지 매출</ThCompactSortable>
                <ThCompactSortable right tone="yellow" sortKey="fullMonthSales" sortConfig={sortConfig} onSort={requestSort}>당월 전체 매출</ThCompactSortable>
                <ThCompactSortable right tone="gray" sortKey="timeGoneGap" sortConfig={sortConfig} onSort={requestSort}>TIME GONE 대비</ThCompactSortable>
                <ThCompactSortable right tone="purple" sortKey="est" sortConfig={sortConfig} onSort={requestSort}>EST</ThCompactSortable>
                <ThCompactSortable right tone="purple" sortKey="estRate" sortConfig={sortConfig} onSort={requestSort}>EST 달성률</ThCompactSortable>
                <ThCompactSortable right tone="green" sortKey="profitAmount" sortConfig={sortConfig} onSort={requestSort}>이익금액</ThCompactSortable>
                <ThCompactSortable right tone="green" sortKey="profitRate" sortConfig={sortConfig} onSort={requestSort}>이익률</ThCompactSortable>
              </tr>
            </thead>
            <tbody>
              {sortedRows.length === 0 ? (
                <tr><td colSpan={compact ? 11 : view === "거래처별" ? 12 : 11} className="border p-8 text-center text-slate-500">표시할 데이터가 없습니다.</td></tr>
              ) : sortedRows.map((r) => (
                <tr key={r.key}>
                  <TdCompact bold>{r.label}</TdCompact>
                  {!compact && view === "거래처별" && (
                    <TdCompact>
                      <div className="font-semibold text-slate-900">{r.lastOrderDate}</div>
                      <div className={`mt-0.5 text-[11px] font-semibold ${r.daysSinceLastOrder >= 30 ? "text-red-600" : r.daysSinceLastOrder >= 7 ? "text-amber-600" : "text-slate-400"}`}>
                        {r.lastOrderDate === "-" ? "발주 없음" : `${r.daysSinceLastOrder}일 경과`}
                      </div>
                    </TdCompact>
                  )}
                  <ClickableAmountCell value={r.prevYearSales} onClick={() => openDrill(r, "prevYear")} />
                  <TdCompact right amount>{pct(r.prevYearRate)}</TdCompact>
                  <ClickableAmountCell value={r.prevMonthSales} onClick={() => openDrill(r, "prevMonth")} />
                  <TdCompact right amount>{pct(r.prevMonthRate)}</TdCompact>
                  <ClickableAmountCell value={r.currentSales} onClick={() => openDrill(r, "current")} />
                  <ClickableAmountCell value={r.fullMonthSales} onClick={() => openDrill(r, "currentFullMonth")} />
                  <TdCompact right amount>{pct(r.timeGoneGap)}</TdCompact>
                  <TdCompact right amount>{won(r.est)}</TdCompact>
                  <TdCompact right amount>{pct(r.estRate)}</TdCompact>
                  <TdCompact right amount>{won(r.profitAmount)}</TdCompact>
                  <TdCompact right amount>{pct(r.profitRate)}</TdCompact>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {drill && <OrderDrillModal title={drill.title} rows={drill.rows} allSales={sales} onClose={() => setDrill(null)} />}
      {inactiveOpen && <InactiveOrdersModal stores={stores} sales={sales} month={month} onClose={() => setInactiveOpen(false)} />}
    </>
  );
}

function ThCompactSortable({ children, sortKey, sortConfig, onSort, right = false, w = "", tone = "default" }: { children: React.ReactNode; sortKey: SalesStatusSortKey; sortConfig: { key: SalesStatusSortKey; direction: SortDirection }; onSort: (key: SalesStatusSortKey) => void; right?: boolean; w?: string; tone?: "default" | "mint" | "blue" | "yellow" | "gray" | "purple" | "green" }) {
  return (
    <ThCompact right={right} w={w} tone={tone}>
      <button type="button" onClick={() => onSort(sortKey)} className={`flex w-full items-center gap-1 ${right ? "justify-end text-right" : "justify-start text-left"}`}>
        <span>{children}</span>
        <span className="text-[10px] text-slate-500">{sortArrow(sortConfig.key === sortKey, sortConfig.direction)}</span>
      </button>
    </ThCompact>
  );
}

function InactiveOrdersModal({ sales, month, onClose }: { stores: Store[]; sales: SalesRecord[]; month: string; onClose: () => void }) {
  const startDate = threeMonthStart(month);
  const endDate = monthEnd(month);

  const itemMap = new Map<string, { itemCode: string; itemName: string; latest?: SalesRecord; recent: boolean }>();
  sales.forEach((r) => {
    const key = `${r.itemCode}|${r.itemName}`;
    const current = itemMap.get(key) || { itemCode: r.itemCode || "-", itemName: r.itemName || "미지정", latest: undefined, recent: false };
    if (inRange(r.saleDate, startDate, endDate)) current.recent = true;
    if (!current.latest || r.saleDate > current.latest.saleDate) current.latest = r;
    itemMap.set(key, current);
  });

  const inactiveItems = Array.from(itemMap.values())
    .filter((item) => !item.recent)
    .map((item) => ({
      itemCode: item.itemCode,
      itemName: item.itemName,
      lastDate: item.latest?.saleDate || "-",
      lastStore: item.latest?.storeName || "-",
      lastAmount: item.latest?.salesAmount || 0,
    }))
    .sort((a, b) => String(a.lastDate).localeCompare(String(b.lastDate)) || a.itemName.localeCompare(b.itemName, "ko-KR"));

  const itemExcelRows = inactiveItems.map((r) => ({
    상품코드: r.itemCode,
    상품명: r.itemName,
    마지막주문일: r.lastDate,
    마지막거래처: r.lastStore,
    마지막매출금액: r.lastAmount,
  }));

  function sendInactiveMail() {
    const to = window.prompt("메일을 받을 주소를 입력하세요. 여러 명이면 쉼표로 구분해주세요.");
    if (!to) return;

    const fileName = `3개월_미주문_품목별_${month}`;
    const subject = `[에이비랩] 3개월 미주문 품목별 현황_${month}`;
    const bodyLines = [
      "안녕하세요.",
      "",
      `${startDate} ~ ${endDate} 기준 3개월 미주문 품목별 현황 공유드립니다.`,
      `총 ${inactiveItems.length.toLocaleString("ko-KR")}건입니다.`,
      "",
      `방금 자동 다운로드된 ${fileName}.xlsx 파일을 첨부해서 전달드립니다.`,
      "",
      "확인 부탁드립니다.",
    ];

    exportExcel(itemExcelRows, fileName);
    window.setTimeout(() => {
      window.location.href = `mailto:${encodeURIComponent(to)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(bodyLines.join("\n"))}`;
    }, 300);
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/55 p-4">
      <div className="flex max-h-[88vh] w-full max-w-5xl flex-col rounded-2xl bg-white shadow-2xl">
        <div className="flex flex-col gap-3 border-b border-slate-200 p-4 lg:flex-row lg:items-center lg:justify-between">
          <div>
            <h3 className="text-lg font-bold text-slate-900">3개월 미주문 품목 현황</h3>
            <p className="mt-1 text-xs text-slate-500">기준 기간: {startDate} ~ {endDate} · 총 {inactiveItems.length.toLocaleString("ko-KR")}건</p>
          </div>
          <div className="flex flex-wrap gap-2">
            <button type="button" onClick={sendInactiveMail} className="rounded-lg bg-blue-600 px-3 py-2 text-xs font-semibold text-white shadow-sm hover:bg-blue-700">
              메일 전송(파일 다운로드)
            </button>
            <button type="button" onClick={() => exportExcel(itemExcelRows, `3개월_미주문_품목별_${month}`)} className="rounded-lg bg-emerald-600 px-3 py-2 text-xs font-semibold text-white shadow-sm hover:bg-emerald-700">
              엑셀 다운로드
            </button>
            <button type="button" onClick={onClose} className="rounded-lg border border-slate-300 bg-white px-3 py-2 text-xs font-semibold text-slate-700 hover:bg-slate-50">
              닫기
            </button>
          </div>
        </div>

        <div className="min-h-0 flex-1 overflow-auto bg-white px-4 pb-4 pt-0">
          <table className="w-full min-w-[860px] border-separate border-spacing-0 border border-slate-200 bg-white text-xs">
            <thead>
              <tr className="bg-slate-100">
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">상품코드</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">상품명</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">마지막 주문일</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">마지막 거래처</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">마지막 매출금액</th>
              </tr>
            </thead>
            <tbody>
              {inactiveItems.length === 0 ? (
                <tr><td colSpan={5} className="border p-8 text-center text-slate-500">3개월 미주문 품목이 없습니다.</td></tr>
              ) : inactiveItems.map((r) => (
                <tr key={`${r.itemCode}|${r.itemName}`} className="hover:bg-slate-50">
                  <td className="border px-2 py-2">{r.itemCode}</td>
                  <td className="border px-2 py-2 font-semibold">{r.itemName}</td>
                  <td className="border px-2 py-2">{r.lastDate}</td>
                  <td className="border px-2 py-2">{r.lastStore}</td>
                  <td className="border px-2 py-2 text-right font-semibold text-slate-900">{won(r.lastAmount)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function DormantAccountPage({ stores, sales, month }: { stores: Store[]; sales: SalesRecord[]; month: string }) {
  const [tab, setTab] = useState<InactiveOrderTab>("거래처별");
  const startDate = threeMonthStart(month);
  const endDate = monthEnd(month);
  const activeStores = stores.filter((s) => s.status === "거래중");
  const recentSales = sales.filter((s) => inRange(s.saleDate, startDate, endDate));
  const recentStoreCodes = new Set(recentSales.map((s) => s.storeCode));
  const latestSaleByStore = new Map<string, SalesRecord>();

  sales.forEach((r) => {
    const prev = latestSaleByStore.get(r.storeCode);
    if (!prev || r.saleDate > prev.saleDate) latestSaleByStore.set(r.storeCode, r);
  });

  const inactiveStores = activeStores
    .filter((s) => !recentStoreCodes.has(s.code))
    .map((s) => {
      const latest = latestSaleByStore.get(s.code);
      return {
        code: s.code,
        name: s.name,
        brand: s.brand,
        manager: s.manager || "미지정",
        channel: s.channel,
        lastDate: latest?.saleDate || "-",
        lastAmount: latest?.salesAmount || 0,
      };
    })
    .sort((a, b) => String(a.lastDate).localeCompare(String(b.lastDate)) || a.name.localeCompare(b.name, "ko-KR"));

  const itemMap = new Map<string, { itemCode: string; itemName: string; latest?: SalesRecord; recent: boolean }>();
  sales.forEach((r) => {
    const key = `${r.itemCode}|${r.itemName}`;
    const current = itemMap.get(key) || { itemCode: r.itemCode || "-", itemName: r.itemName || "미지정", latest: undefined, recent: false };
    if (inRange(r.saleDate, startDate, endDate)) current.recent = true;
    if (!current.latest || r.saleDate > current.latest.saleDate) current.latest = r;
    itemMap.set(key, current);
  });
  const inactiveItems = Array.from(itemMap.values())
    .filter((item) => !item.recent)
    .map((item) => ({
      itemCode: item.itemCode,
      itemName: item.itemName,
      lastDate: item.latest?.saleDate || "-",
      lastStore: item.latest?.storeName || "-",
      lastAmount: item.latest?.salesAmount || 0,
    }))
    .sort((a, b) => String(a.lastDate).localeCompare(String(b.lastDate)) || a.itemName.localeCompare(b.itemName, "ko-KR"));

  const storeExcelRows = inactiveStores.map((r) => ({
    거래처코드: r.code,
    거래처명: r.name,
    브랜드: r.brand,
    담당자: r.manager,
    채널: r.channel,
    마지막주문일: r.lastDate,
    마지막매출금액: r.lastAmount,
  }));
  const itemExcelRows = inactiveItems.map((r) => ({
    상품코드: r.itemCode,
    상품명: r.itemName,
    마지막주문일: r.lastDate,
    마지막거래처: r.lastStore,
    마지막매출금액: r.lastAmount,
  }));

  function sendInactiveMail() {
    const to = window.prompt("메일을 받을 주소를 입력하세요. 여러 명이면 쉼표로 구분해주세요.");
    if (!to) return;

    const targetRows = tab === "거래처별" ? inactiveStores : inactiveItems;
    const excelRows = tab === "거래처별" ? storeExcelRows : itemExcelRows;
    const fileName = `3개월_미주문_${tab}_${month}`;
    const subject = `[에이비랩] 3개월 미주문 ${tab} 현황_${month}`;
    const bodyLines = [
      "안녕하세요.",
      "",
      `${startDate} ~ ${endDate} 기준 3개월 미주문 ${tab} 현황 공유드립니다.`,
      `총 ${targetRows.length.toLocaleString("ko-KR")}건입니다.`,
      "",
      `방금 자동 다운로드된 ${fileName}.xlsx 파일을 첨부해서 전달드립니다.`,
      "",
      "확인 부탁드립니다.",
    ];

    exportExcel(excelRows, fileName);
    window.setTimeout(() => {
      window.location.href = `mailto:${encodeURIComponent(to)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(bodyLines.join("\n"))}`;
    }, 300);
  }

  return (
    <div className="rounded-xl border border-slate-200/70 bg-white/80 p-4 shadow-sm backdrop-blur">
      <div className="mb-3 flex flex-col gap-3 lg:flex-row lg:items-center lg:justify-between">
        <div>
          <h2 className="text-base font-bold text-slate-900">휴면거래처관리</h2>
          <p className="mt-1 text-xs text-slate-500">3개월 미주문 거래처와 품목을 확인하고 엑셀 다운로드 또는 메일 전송용 파일을 만들 수 있습니다. 기준 기간: {startDate} ~ {endDate}</p>
        </div>
        <div className="flex flex-wrap gap-2">
          <button type="button" onClick={sendInactiveMail} className="rounded-lg bg-blue-600 px-3 py-2 text-xs font-semibold text-white shadow-sm hover:bg-blue-700">
            메일 전송용 파일 다운로드
          </button>
          <button type="button" onClick={() => exportExcel(tab === "거래처별" ? storeExcelRows : itemExcelRows, `3개월_미주문_${tab}_${month}`)} className="rounded-lg bg-emerald-600 px-3 py-2 text-xs font-semibold text-white shadow-sm hover:bg-emerald-700">
            엑셀 다운로드
          </button>
        </div>
      </div>

      <div className="mb-3 flex gap-2 border-b border-slate-200 pt-1">
        {(["거래처별", "품목별"] as InactiveOrderTab[]).map((item) => (
          <button key={item} type="button" onClick={() => setTab(item)} className={`rounded-t-lg px-4 py-2 text-sm font-semibold ${tab === item ? "bg-blue-600 text-white" : "bg-slate-100 text-slate-700 hover:bg-slate-200"}`}>
            {item} {item === "거래처별" ? inactiveStores.length : inactiveItems.length}건
          </button>
        ))}
      </div>

      <div className="max-h-[62vh] overflow-auto">
        {tab === "거래처별" ? (
          <table className="w-full min-w-[920px] border border-slate-200 text-xs">
            <thead>
              <tr className="bg-slate-100">
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">거래처코드</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">거래처명</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">브랜드</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">담당자</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">채널</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">마지막 주문일</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">마지막 매출금액</th>
              </tr>
            </thead>
            <tbody>
              {inactiveStores.length === 0 ? (
                <tr><td colSpan={7} className="border p-8 text-center text-slate-500">3개월 미주문 거래처가 없습니다.</td></tr>
              ) : inactiveStores.map((r) => (
                <tr key={r.code} className="hover:bg-slate-50">
                  <td className="border px-2 py-2">{r.code}</td>
                  <td className="border px-2 py-2 font-semibold">{r.name}</td>
                  <td className="border px-2 py-2">{r.brand}</td>
                  <td className="border px-2 py-2">{r.manager}</td>
                  <td className="border px-2 py-2">{r.channel}</td>
                  <td className="border px-2 py-2">{r.lastDate}</td>
                  <td className="border px-2 py-2 text-right font-semibold text-slate-900">{won(r.lastAmount)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        ) : (
          <table className="w-full min-w-[860px] border-separate border-spacing-0 border border-slate-200 bg-white text-xs">
            <thead>
              <tr className="bg-slate-100">
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">상품코드</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">상품명</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">마지막 주문일</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">마지막 거래처</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">마지막 매출금액</th>
              </tr>
            </thead>
            <tbody>
              {inactiveItems.length === 0 ? (
                <tr><td colSpan={5} className="border p-8 text-center text-slate-500">3개월 미주문 품목이 없습니다.</td></tr>
              ) : inactiveItems.map((r) => (
                <tr key={`${r.itemCode}|${r.itemName}`} className="hover:bg-slate-50">
                  <td className="border px-2 py-2">{r.itemCode}</td>
                  <td className="border px-2 py-2 font-semibold">{r.itemName}</td>
                  <td className="border px-2 py-2">{r.lastDate}</td>
                  <td className="border px-2 py-2">{r.lastStore}</td>
                  <td className="border px-2 py-2 text-right font-semibold text-slate-900">{won(r.lastAmount)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}

function ClickableAmountCell({ value, onClick }: { value: number; onClick: () => void }) {
  const disabled = !value;
  return (
    <td className="border border-slate-200 bg-white px-1.5 py-2 text-right align-middle whitespace-normal break-words">
      <button
        type="button"
        onClick={onClick}
        disabled={disabled}
        className={`w-full break-words text-right text-[13px] font-bold leading-snug underline-offset-2 ${disabled ? "cursor-default text-slate-400" : "text-slate-900 hover:underline"}`}
        title={disabled ? "주문내역이 없습니다." : "주문내역 보기"}
      >
        {won(value)}
      </button>
    </td>
  );
}

function compactList(values: string[]) {
  const unique = Array.from(new Set(values.map((v) => norm(v) || "미지정"))).filter(Boolean);
  if (!unique.length) return "-";
  if (unique.length <= 3) return unique.join(", ");
  return `${unique.slice(0, 3).join(", ")} 외 ${unique.length - 3}개`;
}

function OrderDrillModal({ title, rows, allSales, onClose }: { title: string; rows: SalesRecord[]; allSales: SalesRecord[]; onClose: () => void }) {
  const [itemDrill, setItemDrill] = useState<{ itemCode: string; itemName: string; rows: SalesRecord[] } | null>(null);
  const totalSales = sum(rows, "salesAmount");
  const totalQuantity = sum(rows, "quantity");
  const totalProfit = sum(rows, "profitAmount");
  const excelRows = orderRowsForExcel(rows);
  const summaryBrand = compactList(rows.map((r) => r.brand));
  const summaryManager = compactList(rows.map((r) => r.manager || "미지정"));
  const summaryChannel = compactList(rows.map((r) => r.channel));

  const openItemDrill = (itemCode: string, itemName: string) => {
    const filteredRows = allSales
      .filter((r) => r.itemCode === itemCode)
      .sort((a, b) => b.saleDate.localeCompare(a.saleDate) || a.storeName.localeCompare(b.storeName, "ko-KR"));
    setItemDrill({ itemCode, itemName, rows: filteredRows });
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/55 p-3">
      <div className="flex max-h-[94vh] w-full max-w-[96vw] flex-col rounded-2xl bg-white shadow-2xl">
        <div className="sticky top-0 z-20 flex flex-col gap-3 border-b border-slate-200 bg-white p-5 shadow-sm xl:flex-row xl:items-start xl:justify-between">
          <div className="min-w-0 flex-1">
            <div className="flex flex-wrap items-center gap-2">
              <h3 className="text-xl font-bold text-slate-900">{title} 주문내역</h3>
              <span className="rounded-full bg-slate-100 px-3 py-1 text-xs font-semibold text-slate-700">브랜드: {summaryBrand}</span>
              <span className="rounded-full bg-slate-100 px-3 py-1 text-xs font-semibold text-slate-700">담당자: {summaryManager}</span>
              <span className="rounded-full bg-slate-100 px-3 py-1 text-xs font-semibold text-slate-700">채널: {summaryChannel}</span>
            </div>
            <p className="mt-2 text-sm text-slate-500">총 {rows.length.toLocaleString("ko-KR")}건 · 수량 {won(totalQuantity)} · 매출 {won(totalSales)}원 · 이익 {won(totalProfit)}원</p>
            <p className="mt-1 text-xs text-slate-400">상품코드를 클릭하면 해당 품목이 나간 전체 거래처 주문내역을 볼 수 있습니다.</p>
          </div>
          <div className="flex shrink-0 gap-2">
            <button
              type="button"
              onClick={() => exportExcel(excelRows, `${title.replaceAll(" ", "_").replaceAll("·", "_")}_주문내역`)}
              className="rounded-lg bg-emerald-600 px-4 py-2 text-sm font-semibold text-white shadow-sm hover:bg-emerald-700"
            >
              주문내역 엑셀 다운로드
            </button>
            <button
              type="button"
              onClick={onClose}
              className="rounded-lg border border-slate-300 bg-white px-4 py-2 text-sm font-semibold text-slate-700 hover:bg-slate-50"
            >
              닫기
            </button>
          </div>
        </div>

        <div className="relative isolate min-h-0 flex-1 overflow-auto bg-white px-5 pb-5 pt-0">
          <table className="w-full min-w-[1100px] border-separate border-spacing-0 border border-slate-200 bg-white text-sm">
            <thead className="sticky top-0 z-[80] bg-white shadow-[0_2px_0_0_#e2e8f0]">
              <tr className="bg-white">
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-3 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">주문일</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-3 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">거래처</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-3 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">상품코드</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-3 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">상품명</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-3 py-2 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">수량</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-3 py-2 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">매출금액</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-3 py-2 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">원가금액</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-3 py-2 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">이익금액</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-3 py-2 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">이익률</th>
              </tr>
            </thead>
            <tbody>
              {rows.length === 0 ? (
                <tr><td colSpan={9} className="border p-10 text-center text-slate-500">표시할 주문내역이 없습니다.</td></tr>
              ) : rows.map((r) => (
                <tr key={r.id} className="hover:bg-slate-50">
                  <td className="border px-3 py-2">{r.saleDate}</td>
                  <td className="border px-3 py-2 font-semibold">{r.storeName}</td>
                  <td className="border px-3 py-2">
                    <button
                      type="button"
                      onClick={() => openItemDrill(r.itemCode, r.itemName)}
                      className="font-bold text-slate-900 underline-offset-2 hover:underline"
                      title="품목별 주문내역 보기"
                    >
                      {r.itemCode}
                    </button>
                  </td>
                  <td className="border px-3 py-2">{r.itemName}</td>
                  <td className="border px-3 py-2 text-right">{won(r.quantity)}</td>
                  <td className="border px-3 py-2 text-right text-base font-bold text-slate-900">{won(r.salesAmount)}</td>
                  <td className="border px-3 py-2 text-right font-semibold text-slate-900">{won(r.costAmount)}</td>
                  <td className="border px-3 py-2 text-right font-semibold text-slate-900">{won(r.profitAmount)}</td>
                  <td className="border px-3 py-2 text-right text-slate-900">{pct(r.profitRate)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {itemDrill && (
        <ItemDrillModal
          itemCode={itemDrill.itemCode}
          itemName={itemDrill.itemName}
          rows={itemDrill.rows}
          onClose={() => setItemDrill(null)}
        />
      )}
    </div>
  );
}

function ItemDrillModal({ itemCode, itemName, rows, onClose }: { itemCode: string; itemName: string; rows: SalesRecord[]; onClose: () => void }) {
  const totalSales = sum(rows, "salesAmount");
  const totalQuantity = sum(rows, "quantity");
  const totalCost = sum(rows, "costAmount");
  const totalProfit = sum(rows, "profitAmount");
  const totalProfitRate = weightedProfitRate(rows);
  const excelRows = orderRowsForExcel(rows);

  return (
    <div className="absolute inset-0 z-[60] flex items-center justify-center bg-slate-950/45 p-4">
      <div className="flex max-h-[88vh] w-full max-w-[90vw] flex-col rounded-2xl bg-white shadow-2xl ring-1 ring-slate-200">
        <div className="sticky top-0 z-20 flex flex-col gap-3 border-b border-slate-200 bg-white p-5 shadow-sm xl:flex-row xl:items-start xl:justify-between">
          <div className="min-w-0 flex-1">
            <h3 className="text-xl font-bold text-slate-900">품목별 전체 거래처 주문내역</h3>
            <p className="mt-1 text-sm font-semibold text-slate-700">{itemCode} · {itemName}</p>
            <p className="mt-2 text-sm text-slate-500">총 {rows.length.toLocaleString("ko-KR")}건 · 수량 {won(totalQuantity)} · 매출 {won(totalSales)}원 · 원가 {won(totalCost)}원 · 이익 {won(totalProfit)}원 · 이익률 {pct(totalProfitRate)}</p>
          </div>
          <div className="flex shrink-0 gap-2">
            <button
              type="button"
              onClick={() => exportExcel(excelRows, `${itemCode}_${itemName}_품목별_전체거래처_주문내역`)}
              className="rounded-lg bg-emerald-600 px-4 py-2 text-sm font-semibold text-white shadow-sm hover:bg-emerald-700"
            >
              품목별 전체거래처 엑셀 다운로드
            </button>
            <button
              type="button"
              onClick={onClose}
              className="rounded-lg border border-slate-300 bg-white px-4 py-2 text-sm font-semibold text-slate-700 hover:bg-slate-50"
            >
              닫기
            </button>
          </div>
        </div>

        <div className="relative isolate min-h-0 flex-1 overflow-auto bg-white px-5 pb-5 pt-0">
          <table className="w-full min-w-[920px] border-separate border-spacing-0 border border-slate-200 bg-white text-sm">
            <thead className="sticky top-0 z-[80] bg-white shadow-[0_2px_0_0_#e2e8f0]">
              <tr className="bg-white">
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-3 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">주문일</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-3 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">거래처</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-3 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">담당자</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-3 py-2 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">수량</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-3 py-2 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">매출금액</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-3 py-2 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">원가금액</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-3 py-2 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">이익금액</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-3 py-2 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">이익률</th>
              </tr>
            </thead>
            <tbody>
              {rows.length === 0 ? (
                <tr><td colSpan={8} className="border p-10 text-center text-slate-500">표시할 품목별 주문내역이 없습니다.</td></tr>
              ) : rows.map((r) => (
                <tr key={r.id} className="hover:bg-slate-50">
                  <td className="border px-3 py-2">{r.saleDate}</td>
                  <td className="border px-3 py-2 font-semibold">{r.storeName}</td>
                  <td className="border px-3 py-2">{r.manager || "미지정"}</td>
                  <td className="border px-3 py-2 text-right">{won(r.quantity)}</td>
                  <td className="border px-3 py-2 text-right text-base font-bold text-slate-900">{won(r.salesAmount)}</td>
                  <td className="border px-3 py-2 text-right font-semibold text-slate-900">{won(r.costAmount)}</td>
                  <td className="border px-3 py-2 text-right font-semibold text-slate-900">{won(r.profitAmount)}</td>
                  <td className="border px-3 py-2 text-right text-slate-900">{pct(r.profitRate)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function FilterAmountCard({ title, value, tone, onClick }: { title: string; value: number | string; tone: "mint" | "blue" | "yellow" | "green"; onClick?: () => void }) {
  const toneClass =
    tone === "mint" ? "border-emerald-200 bg-emerald-50" :
    tone === "blue" ? "border-blue-200 bg-blue-50" :
    tone === "yellow" ? "border-amber-200 bg-amber-50" :
    "border-green-200 bg-green-50";

  return (
    <button
      type="button"
      onClick={onClick}
      disabled={!onClick}
      className={`rounded-xl border px-3 py-2 text-left shadow-sm ${toneClass} ${onClick ? "cursor-pointer hover:shadow-md" : "cursor-default"}`}
    >
      <p className="text-[11px] font-semibold text-slate-500">{title}</p>
      <p className="mt-1 break-all text-right text-base font-bold text-slate-900">{typeof value === "number" ? won(value) : value}</p>
    </button>
  );
}

function ThCompact({ children, right = false, w = "", tone = "default" }: { children: React.ReactNode; right?: boolean; w?: string; tone?: "default" | "mint" | "blue" | "yellow" | "gray" | "purple" | "green" }) {
  const toneClass =
    tone === "mint" ? "border-emerald-200 bg-emerald-50 text-emerald-800" :
    tone === "blue" ? "border-blue-200 bg-blue-50 text-blue-800" :
    tone === "yellow" ? "border-amber-200 bg-amber-50 text-amber-800" :
    tone === "gray" ? "border-slate-300 bg-slate-100 text-slate-800" :
    tone === "purple" ? "border-violet-200 bg-violet-50 text-violet-800" :
    tone === "green" ? "border-green-200 bg-green-50 text-green-800" :
    "border-slate-200 bg-slate-100 text-slate-800";

  return (
    <th className={`sticky top-0 z-50 border px-1 py-1.5 align-middle text-[11px] font-bold leading-tight whitespace-normal break-keep shadow-sm bg-clip-padding ${toneClass} ${right ? "text-right" : "text-left"} ${w}`}>
      {children}
    </th>
  );
}

function TdCompact({ children, right = false, bold = false, color = "", amount = false }: { children: React.ReactNode; right?: boolean; bold?: boolean; color?: string; amount?: boolean }) {
  return <td className={`border border-slate-200 bg-white px-1.5 py-2 align-middle whitespace-normal break-words ${right ? "text-right" : "text-left"} ${bold ? "font-semibold" : ""} ${amount ? "text-[13px] font-bold leading-snug text-slate-900" : ""} ${color}`}>{children}</td>;
}

function SalesCompare({ stores, sales, month, date }: { stores: Store[]; sales: SalesRecord[]; month: string; date: string }) {
  const [search, setSearch] = useState("");
  const [searchDraft, setSearchDraft] = useState("");
  const stMap = storeMap(stores);
  const storeCodes = Array.from(new Set([...stores.map((s) => s.code), ...sales.map((s) => s.storeCode)]));
  const rows = storeCodes.map((code) => {
    const store = stMap.get(code);
    const current = sales.filter((s) => s.period === "current" && s.storeCode === code && inRange(s.saleDate, monthStart(month), date));
    const prevMonth = sales.filter((s) => s.period === "prevMonth" && s.refMonth === month && s.storeCode === code);
    const prevYear = sales.filter((s) => s.period === "prevYear" && s.refMonth === month && s.storeCode === code);
    const currentSales = sum(current, "salesAmount");
    const prevMonthSales = sum(prevMonth, "salesAmount");
    const prevYearSales = sum(prevYear, "salesAmount");
    return {
      code,
      name: store?.name || current[0]?.storeName || prevMonth[0]?.storeName || prevYear[0]?.storeName || "",
      currentSales,
      prevMonthSales,
      prevYearSales,
      prevMonthRate: prevMonthSales ? ((currentSales - prevMonthSales) / prevMonthSales) * 100 : 0,
      prevYearRate: prevYearSales ? ((currentSales - prevYearSales) / prevYearSales) * 100 : 0,
    };
  }).filter((r) => `${r.code} ${r.name}`.toLowerCase().includes(search.toLowerCase()));

  const salesCompareExcelRows = rows.map((r) => ({
    거래처코드: r.code,
    거래처명: r.name,
    전년동월: r.prevYearSales,
    전월: r.prevMonthSales,
    당월: r.currentSales,
    "전년동월 대비": pct(r.prevYearRate),
    "전월 대비": pct(r.prevMonthRate),
  }));

  return (
    <div className="rounded-xl border border-slate-200/70 bg-white/80 p-4 shadow-sm backdrop-blur">
      <div className="mb-3 flex flex-col gap-2 lg:flex-row lg:items-center lg:justify-between">
        <div>
          <h2 className="text-lg font-bold text-slate-900">매출비교</h2>
          <p className="mt-1 text-sm text-slate-500">거래처별 전년동월 / 전월 / 당월 매출을 비교합니다.</p>
        </div>
        <div className="flex flex-wrap items-center gap-2">
          <input
            value={searchDraft}
            onChange={(e) => {
              const value = e.target.value;
              setSearchDraft(value);
              if (!value.trim()) setSearch("");
            }}
            onKeyDown={(e) => {
              if (e.key === "Enter") setSearch(searchDraft.trim());
            }}
            placeholder="거래처코드 / 거래처명 검색"
            className="w-[300px] rounded-xl border border-slate-300 bg-white/80 px-3 py-2 text-sm outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
          />
          <button
            onClick={() => exportExcel(salesCompareExcelRows, `매출비교_${month}`)}
            className="rounded-xl bg-emerald-600 px-4 py-2 text-sm font-semibold text-white shadow-sm hover:bg-emerald-700"
          >
            매출비교 엑셀 다운로드
          </button>
        </div>
      </div>

      <div className="max-h-[68vh] overflow-visible">
        <table className="w-full table-fixed border border-slate-200 text-xs">
          <thead>
            <tr className="bg-slate-100">
              <Th>거래처코드</Th>
              <Th>거래처명</Th>
              <Th right>전년동월</Th>
              <Th right>전월</Th>
              <Th right>당월</Th>
              <Th right>전년동월 대비</Th>
              <Th right>전월 대비</Th>
            </tr>
          </thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r.code}>
                <Td>{r.code}</Td>
                <Td bold>{r.name}</Td>
                <Td right>{won(r.prevYearSales)}</Td>
                <Td right>{won(r.prevMonthSales)}</Td>
                <Td right>{won(r.currentSales)}</Td>
                <Td right>{pct(r.prevYearRate)}</Td>
                <Td right>{pct(r.prevMonthRate)}</Td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function MonthStartManagement({
  stores,
  setStores,
  sales,
  setSales,
  targets,
  setTargets,
  ests,
  setEsts,
  month,
  date,
  timeConfigs,
  setTimeConfigs,
  codeMappings,
  setCodeMappings,
}: {
  stores: Store[];
  setStores: (v: Store[]) => void;
  sales: SalesRecord[];
  setSales: React.Dispatch<React.SetStateAction<SalesRecord[]>>;
  targets: TargetRecord[];
  setTargets: (v: TargetRecord[]) => void;
  ests: EstRecord[];
  setEsts: (v: EstRecord[]) => void;
  month: string;
  date: string;
  timeConfigs: TimeConfig[];
  setTimeConfigs: (v: TimeConfig[]) => void;
  codeMappings: StoreCodeMapping[];
  setCodeMappings: (v: StoreCodeMapping[]) => void;
}) {
  const [tab, setTab] = useState<MonthStartTab>("거래처/휴면 관리");

  return (
    <div className="space-y-4">
      <div className="rounded-2xl border border-slate-200/70 bg-white/80 p-4 shadow-sm backdrop-blur">
        <div className="flex flex-wrap gap-2">
          {MONTH_TABS.map((t) => (
            <button key={t} onClick={() => setTab(t)} className={`rounded-xl px-4 py-2 text-sm font-semibold ${tab === t ? "bg-blue-600 text-white" : "bg-slate-100 text-slate-700 hover:bg-slate-200"}`}>
              {t}
            </button>
          ))}
        </div>
      </div>

      {tab === "거래처/휴면 관리" && (
        <div className="space-y-4">
          <MappingPage stores={stores} setStores={setStores} sales={sales} month={month} codeMappings={codeMappings} setCodeMappings={setCodeMappings} />
        </div>
      )}
      {tab === "Target/EST 관리" && (
        <div className="space-y-4">
          <TargetByTypePage records={targets} setRecords={setTargets} month={month} />
          <TargetOrEstPage title="EST 관리" records={ests} setRecords={setEsts} stores={stores} month={month} />
        </div>
      )}
      {tab === "업로드 관리" && <UploadPage stores={stores} setStores={setStores} sales={sales} setSales={setSales} month={month} date={date} timeConfigs={timeConfigs} setTimeConfigs={setTimeConfigs} />}
    </div>
  );
}

function MappingPage({ stores, setStores, sales, month, codeMappings, setCodeMappings }: { stores: Store[]; setStores: (v: Store[]) => void; sales: SalesRecord[]; month: string; codeMappings: StoreCodeMapping[]; setCodeMappings: (v: StoreCodeMapping[]) => void }) {
  const empty: Store = { code: "", name: "", channel: "도매", manager: "", storeType: "비매장", brand: "미지정", status: "거래중" };
  const [form, setForm] = useState<Store>(empty);
  const [search, setSearch] = useState("");
  const [searchDraft, setSearchDraft] = useState("");
  const [statusFilter, setStatusFilter] = useState<"all" | "active" | "inactive">("active");
  const [mappingListOpen, setMappingListOpen] = useState(false);
  const [mappingModalOpen, setMappingModalOpen] = useState(false);
  const rows = stores
    .filter((s) =>
      statusFilter === "all" ||
      (statusFilter === "active" && s.status === "거래중") ||
      (statusFilter === "inactive" && s.status === "거래종료")
    )
    .filter((s) => `${s.code} ${s.name} ${s.channel} ${s.manager} ${s.brand}`.toLowerCase().includes(search.toLowerCase()));
  const activeCount = stores.filter((s) => s.status === "거래중").length;
  const inactiveCount = stores.filter((s) => s.status === "거래종료").length;

  function save() {
    if (!form.code || !form.name) return alert("거래처코드와 거래처명은 필수입니다.");
    const exists = stores.some((s) => s.code === form.code);
    setStores(exists ? stores.map((s) => s.code === form.code ? form : s) : [...stores, form]);
    setForm(empty);
  }

  function toggleStatus(store: Store, e: React.MouseEvent<HTMLButtonElement>) {
    e.stopPropagation();
    const nextStatus: Store["status"] = store.status === "거래중" ? "거래종료" : "거래중";
    setStores(stores.map((s) => s.code === store.code ? { ...s, status: nextStatus } : s));
    if (form.code === store.code) setForm({ ...store, status: nextStatus });
  }

  function deleteStore(store: Store, e: React.MouseEvent<HTMLButtonElement>) {
    e.stopPropagation();
    const usedCount = sales.filter((r) => r.storeCode === store.code).length;
    const warning = usedCount
      ? `\n\n이 거래처는 매출 데이터 ${usedCount.toLocaleString("ko-KR")}건에 사용되고 있습니다. 삭제해도 기존 매출 데이터는 남아 있지만, 거래처 매핑 정보는 삭제됩니다.`
      : "";
    if (!confirm(`${store.name} 거래처를 삭제할까요?${warning}`)) return;
    setStores(stores.filter((s) => s.code !== store.code));
    if (form.code === store.code) setForm(empty);
  }

  async function upload(file: File | null) {
    if (!file) return;
    const rows = await readFileRows(file);
    const parsed: Store[] = rows.map((r) => {
      const channel = normalizeChannel(r["채널1"] ?? r["채널"]);
      return {
        code: norm(r["거래처코드"] ?? r["거래처 코드"] ?? r["매장코드"]),
        name: norm(r["거래처명"] ?? r["매장명"]),
        channel,
        manager: norm(r["담당자"]) as Manager,
        storeType: normalizeStoreType(r["채널2"] ?? r["매장구분"], channel),
        brand: norm(r["브랜드"]) || "미지정",
        status: normalizeStatus(r["거래상태"]),
      };
    }).filter((s) => s.code && s.name);

    const map = new Map(stores.map((s) => [s.code, s]));
    parsed.forEach((s) => map.set(s.code, { ...map.get(s.code), ...s }));
    setStores(Array.from(map.values()));
    alert(`거래처 매핑 ${parsed.length}건을 반영했습니다.`);
  }


  const currentSalesStores = useMemo(() => {
    const map = new Map<string, { code: string; name: string }>();
    sales
      .filter((r) => r.period === "current" && inRange(r.saleDate, monthStart(month), monthEnd(month)))
      .forEach((r) => {
        if (!map.has(r.storeCode)) map.set(r.storeCode, { code: r.storeCode, name: r.storeName });
      });
    return Array.from(map.values()).sort((a, b) => a.name.localeCompare(b.name, "ko-KR"));
  }, [sales, month]);

  const mappingCheckRows = useMemo(() => {
    const currentByCode = new Map(currentSalesStores.map((s) => [s.code, s]));
    const currentByName = new Map<string, { code: string; name: string }>();
    currentSalesStores.forEach((s) => {
      const key = normalizeStoreNameKey(s.name);
      if (key && !currentByName.has(key)) currentByName.set(key, s);
    });

    const sourceMap = new Map<string, { period: PeriodType; code: string; name: string; amount: number }>();
    sales
      .filter((r) => r.period === "prevYear" && r.refMonth === month)
      .forEach((r) => {
        const key = `${r.period}|${r.storeCode}|${r.storeName}`;
        const item = sourceMap.get(key) || { period: r.period, code: r.storeCode, name: r.storeName, amount: 0 };
        item.amount += r.salesAmount;
        sourceMap.set(key, item);
      });

    return Array.from(sourceMap.values()).map((r) => {
      const byCode = currentByCode.get(r.code);
      const byName = currentByName.get(normalizeStoreNameKey(r.name));
      const manual = codeMappings.find((m) => norm(m.oldCode) === norm(r.code) && (!m.oldName || normalizeStoreNameKey(m.oldName) === normalizeStoreNameKey(r.name)));
      let category = "수동 매핑 필요";
      let reason = "당월 매출이 없어 전년동월 업로드 거래처 기준으로 표시";
      let targetCode = "";
      let targetName = "";

      if (manual) {
        category = "수동 매핑 완료";
        reason = "사용자가 현재 거래처로 직접 매핑함";
        targetCode = manual.currentCode;
        targetName = manual.currentName;
      } else if (byCode && normalizeStoreNameKey(byCode.name) === normalizeStoreNameKey(r.name)) {
        category = "자동 매핑";
        reason = "거래처명과 거래처코드가 모두 당월과 같음";
        targetCode = byCode.code;
        targetName = byCode.name;
      } else if (byCode) {
        category = "자동 매핑";
        reason = "거래처코드가 당월과 같음";
        targetCode = byCode.code;
        targetName = byCode.name;
      } else if (byName) {
        category = "자동 매핑";
        reason = "거래처명이 당월과 같음";
        targetCode = byName.code;
        targetName = byName.name;
      }

      return { ...r, category, reason, targetCode, targetName };
    }).sort((a, b) => a.category.localeCompare(b.category, "ko-KR") || a.name.localeCompare(b.name, "ko-KR"));
  }, [sales, month, currentSalesStores, codeMappings]);

  const mappingSummary = useMemo(() => {
    return mappingCheckRows.reduce((acc, row) => {
      acc[row.category] = (acc[row.category] || 0) + 1;
      return acc;
    }, {} as Record<string, number>);
  }, [mappingCheckRows]);

  function saveManualMapping(row: { code: string; name: string; targetCode: string; targetName: string }) {
    const currentCode = window.prompt("현재 거래처코드를 입력하세요.", row.targetCode || "");
    if (!currentCode) return;
    const matched = currentSalesStores.find((s) => s.code === currentCode) || stores.find((s) => s.code === currentCode);
    const currentName = window.prompt("현재 거래처명을 입력하세요.", row.targetName || matched?.name || "");
    if (!currentName) return;
    const next: StoreCodeMapping = {
      id: `${row.code}|${row.name}|${currentCode}`,
      oldCode: row.code,
      oldName: row.name,
      currentCode,
      currentName,
    };
    setCodeMappings([...codeMappings.filter((m) => !(norm(m.oldCode) === norm(row.code) && normalizeStoreNameKey(m.oldName) === normalizeStoreNameKey(row.name))), next]);
  }

  function deleteManualMapping(row: { code: string; name: string }) {
    setCodeMappings(codeMappings.filter((m) => !(norm(m.oldCode) === norm(row.code) && normalizeStoreNameKey(m.oldName) === normalizeStoreNameKey(row.name))));
  }

  return (
    <div className="rounded-xl border border-slate-200/70 bg-white/80 p-4 shadow-sm backdrop-blur">
      <div className="mb-3 flex flex-col gap-2 lg:flex-row lg:items-center lg:justify-between">
        <div>
          <h2 className="text-base font-bold">거래처 매핑관리</h2>
          <p className="mt-1 text-xs text-slate-500">채널/담당자/브랜드 정보를 거래처코드 기준으로 관리합니다. 전년동월 매출은 당월 매출에 같은 거래처코드나 거래처명이 있으면 당월 거래처 기준으로 합산하고, 당월 매출이 없으면 전년동월 업로드 거래처 기준으로 별도 표시합니다.</p>
        </div>
        <label className="shrink-0 cursor-pointer rounded-md bg-green-600 px-2 py-1 text-[11px] font-semibold text-white">
          매핑 엑셀 업로드
          <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={(e) => upload(e.target.files?.[0] || null)} />
        </label>
      </div>

      <div className="sticky top-0 z-20 mb-3 rounded-xl border border-blue-100 bg-blue-50/80 p-2 shadow-sm backdrop-blur">
        <div className="flex flex-wrap items-center gap-2">
          <input
            value={searchDraft}
            onChange={(e) => {
              const value = e.target.value;
              setSearchDraft(value);
              if (!value.trim()) setSearch("");
            }}
            onKeyDown={(e) => {
              if (e.key === "Enter") {
                setSearch(searchDraft.trim());
                setMappingListOpen(true);
              }
            }}
            placeholder="거래처 검색"
            className="h-8 w-[260px] rounded-lg border border-slate-300 bg-white px-2 py-1 text-xs outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
          />
          <div className="flex overflow-hidden rounded-lg border border-slate-300 bg-white text-xs font-semibold shadow-sm">
            <button type="button" onClick={() => { setStatusFilter("all"); setMappingListOpen(true); }} className={`px-3 py-1.5 ${statusFilter === "all" ? "bg-slate-900 text-white" : "text-slate-600 hover:bg-slate-50"}`}>전체 {stores.length.toLocaleString("ko-KR")}</button>
            <button type="button" onClick={() => { setStatusFilter("active"); setMappingListOpen(true); }} className={`border-l border-slate-200 px-3 py-1.5 ${statusFilter === "active" ? "bg-emerald-600 text-white" : "text-slate-600 hover:bg-slate-50"}`}>활성 {activeCount.toLocaleString("ko-KR")}</button>
            <button type="button" onClick={() => { setStatusFilter("inactive"); setMappingListOpen(true); }} className={`border-l border-slate-200 px-3 py-1.5 ${statusFilter === "inactive" ? "bg-slate-600 text-white" : "text-slate-600 hover:bg-slate-50"}`}>비활성 {inactiveCount.toLocaleString("ko-KR")}</button>
          </div>
          <button
            type="button"
            onClick={() => setMappingListOpen((v) => !v)}
            className="rounded-lg border border-slate-300 bg-white px-3 py-1.5 text-xs font-semibold text-slate-700 shadow-sm hover:bg-slate-50"
          >
            {mappingListOpen ? "거래처 목록 닫기" : "거래처 목록 열기"}
          </button>
          <button
            type="button"
            onClick={() => setMappingModalOpen(true)}
            className="rounded-lg bg-orange-500 px-3 py-1.5 text-xs font-semibold text-white shadow-sm hover:bg-orange-600"
          >
            검증표 열기
          </button>
          <span className="text-xs font-semibold text-slate-500">표시 {rows.length.toLocaleString("ko-KR")}건</span>
        </div>
      </div>

      {mappingListOpen && (
      <div className="mb-3 rounded-xl border border-slate-200 bg-slate-50 p-3">
        <div className="mb-2 text-xs font-bold text-slate-700">신규 매장 추가 / 거래처 수정</div>
        <div className="flex flex-wrap items-center gap-1.5">
          <input value={form.brand} onChange={(e) => setForm({ ...form, brand: e.target.value })} placeholder="브랜드" className="h-8 w-[150px] rounded-md border bg-white px-2 py-1 text-xs" />
          <input value={form.code} onChange={(e) => setForm({ ...form, code: e.target.value })} placeholder="거래처코드" className="h-8 w-[140px] rounded-md border bg-white px-2 py-1 text-xs" />
          <input value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} placeholder="거래처명" className="h-8 w-[220px] rounded-md border bg-white px-2 py-1 text-xs" />
          <select value={form.channel} onChange={(e) => setForm({ ...form, channel: e.target.value as Channel })} className="h-8 w-[110px] rounded-md border bg-white px-2 py-1 text-xs">
            <option value="">채널1</option>
            {CHANNELS.map((c) => <option key={c}>{c}</option>)}
          </select>
          <select value={form.storeType} onChange={(e) => setForm({ ...form, storeType: e.target.value as StoreType })} className="h-8 w-[105px] rounded-md border bg-white px-2 py-1 text-xs">
            <option>매장</option>
            <option>비매장</option>
          </select>
          <select value={form.manager} onChange={(e) => setForm({ ...form, manager: e.target.value as Manager })} className="h-8 w-[100px] rounded-md border bg-white px-2 py-1 text-xs">
            <option value="">담당자</option>
            {MANAGERS.map((m) => <option key={m}>{m}</option>)}
          </select>
          <button onClick={save} className="h-8 w-[70px] rounded-md bg-blue-600 px-2 py-1 text-xs font-semibold text-white">저장</button>
          <button type="button" onClick={() => setForm(empty)} className="h-8 rounded-md border border-slate-300 bg-white px-3 py-1 text-xs font-semibold text-slate-700">초기화</button>
          <span className="text-[11px] font-medium text-slate-500">신규 추가 시 상태는 자동으로 활성 처리됩니다. 기존 거래처 수정 시 상태는 유지됩니다.</span>
        </div>
      </div>
      )}

      {mappingModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/55 p-4">
          <div className="flex max-h-[92vh] w-full max-w-[96vw] flex-col rounded-2xl bg-white shadow-2xl">
            <div className="flex items-center justify-between border-b border-slate-200 p-4">
              <div>
                <h3 className="text-lg font-bold text-slate-900">전년동월 거래처 매핑 검증</h3>
                <p className="mt-1 text-xs text-slate-500">검증표는 필요할 때만 열어서 수동 매핑을 설정합니다.</p>
              </div>
              <button type="button" onClick={() => setMappingModalOpen(false)} className="rounded-lg border border-slate-300 bg-white px-3 py-2 text-xs font-semibold text-slate-700 hover:bg-slate-50">닫기</button>
            </div>
            <div className="min-h-0 flex-1 overflow-auto bg-white px-4 pb-4 pt-0">
            <div className="rounded-xl border border-slate-200 bg-slate-50 p-3">
        <div className="mb-2 flex flex-col gap-1 lg:flex-row lg:items-center lg:justify-between">
          <div>
            <h3 className="text-sm font-bold text-slate-900">전년동월 거래처 매핑 검증</h3>
            <p className="mt-1 text-xs text-slate-500">전년동월 업로드 거래처를 당월 매출 거래처 기준으로 비교합니다. 코드나 거래처명 중 하나라도 당월 매출과 같으면 당월 거래처 기준으로 자동 합산되고, 당월 매출이 없으면 전년동월 업로드 거래처 기준으로 표시됩니다.</p>
          </div>
          <div className="flex flex-wrap gap-2 text-xs font-semibold text-slate-600">
            <span>자동 {mappingSummary["자동 매핑"] || 0}건</span>
            <span>수동필요 {mappingSummary["수동 매핑 필요"] || 0}건</span>
            <span>수동완료 {mappingSummary["수동 매핑 완료"] || 0}건</span>
          </div>
        </div>
        <div className="max-h-[72vh] overflow-auto">
          <table className="w-full min-w-[960px] border-separate border-spacing-0 border border-slate-200 bg-white text-xs">
            <thead>
              <tr className="bg-slate-100">
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-1.5 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">구분</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-1.5 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">업로드구분</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-1.5 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">업로드코드</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-1.5 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">업로드거래처명</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-1.5 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">업로드매출금액</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-1.5 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">당월코드</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-1.5 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">당월거래처명</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-1.5 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">매핑방식</th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-1.5 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">관리</th>
              </tr>
            </thead>
            <tbody>
              {mappingCheckRows.length === 0 ? (
                <tr><td colSpan={9} className="border p-5 text-center text-slate-500">검증할 전년동월 매출 데이터가 없습니다.</td></tr>
              ) : mappingCheckRows.map((r) => (
                <tr key={`${r.period}|${r.code}|${r.name}`} className="hover:bg-slate-50">
                  <td className="border px-2 py-1.5 font-semibold">{r.category}</td>
                  <td className="border px-2 py-1.5">전년동월</td>
                  <td className="border px-2 py-1.5">{r.code}</td>
                  <td className="border px-2 py-1.5 font-semibold">{r.name}</td>
                  <td className="border px-2 py-1.5 text-right font-semibold">{won(r.amount)}</td>
                  <td className="border px-2 py-1.5">{r.targetCode || "-"}</td>
                  <td className="border px-2 py-1.5">{r.targetName || "-"}</td>
                  <td className="border px-2 py-1.5">{r.reason}</td>
                  <td className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-1.5 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">
                    {r.category === "수동 매핑 완료" ? (
                      <button type="button" onClick={() => deleteManualMapping(r)} className="rounded-md border border-red-200 bg-red-50 px-2 py-0.5 text-[11px] font-bold text-red-600 hover:bg-red-100">해제</button>
                    ) : r.category === "수동 매핑 필요" ? (
                      <button type="button" onClick={() => saveManualMapping(r)} className="rounded-md bg-blue-600 px-2 py-0.5 text-[11px] font-bold text-white hover:bg-blue-700">수동매핑</button>
                    ) : (
                      <span className="text-slate-400">자동</span>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

            </div>
          </div>
        </div>
      )}

      {mappingListOpen && (
      <div className="max-h-[60vh] overflow-auto">
        <table className="w-auto table-fixed border text-xs">
          <colgroup>
            <col className="w-[130px]" />
            <col className="w-[200px]" />
            <col className="w-[92px]" />
            <col className="w-[86px]" />
            <col className="w-[86px]" />
            <col className="w-[120px]" />
            <col className="w-[78px]" />
            <col className="w-[56px]" />
          </colgroup>
          <thead>
            <tr className="bg-slate-100">
              <th className="border border-slate-200 bg-slate-100 px-2 py-1.5 text-left text-xs font-bold">거래처코드</th>
              <th className="border border-slate-200 bg-slate-100 px-2 py-1.5 text-left text-xs font-bold">거래처명</th>
              <th className="border border-slate-200 bg-slate-100 px-2 py-1.5 text-left text-xs font-bold">채널</th>
              <th className="border border-slate-200 bg-slate-100 px-2 py-1.5 text-left text-xs font-bold">담당자</th>
              <th className="border border-slate-200 bg-slate-100 px-2 py-1.5 text-left text-xs font-bold">매장구분</th>
              <th className="border border-slate-200 bg-slate-100 px-2 py-1.5 text-left text-xs font-bold">브랜드</th>
              <th className="border border-slate-200 bg-slate-100 px-2 py-1.5 text-left text-xs font-bold">상태</th>
              <th className="border border-slate-200 bg-slate-100 px-2 py-1.5 text-right text-xs font-bold">관리</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((s) => (
              <tr key={s.code} onClick={() => setForm(s)} className="cursor-pointer hover:bg-blue-50">
                <td className="truncate border border-slate-200 bg-white px-2 py-1.5 text-left align-middle" title={s.code}>{s.code}</td>
                <td className="truncate border border-slate-200 bg-white px-2 py-1.5 text-left align-middle font-semibold" title={s.name}>{s.name}</td>
                <td className="border border-slate-200 bg-white px-2 py-1.5 text-left align-middle">{s.channel}</td>
                <td className="border border-slate-200 bg-white px-2 py-1.5 text-left align-middle">{s.manager || "-"}</td>
                <td className="border border-slate-200 bg-white px-2 py-1.5 text-left align-middle">{s.storeType}</td>
                <td className="truncate border border-slate-200 bg-white px-2 py-1.5 text-left align-middle" title={s.brand}>{s.brand}</td>
                <td className="border border-slate-200 bg-white px-2 py-1.5 text-left align-middle">
                  <button
                    type="button"
                    onClick={(e) => toggleStatus(s, e)}
                    className={`rounded-full px-2 py-0.5 text-[11px] font-semibold ${s.status === "거래중" ? "bg-emerald-50 text-emerald-700 ring-1 ring-emerald-200" : "bg-slate-100 text-slate-500 ring-1 ring-slate-200"}`}
                    title={s.status === "거래중" ? "클릭하면 비활성 처리됩니다." : "클릭하면 다시 활성화됩니다."}
                  >
                    {s.status === "거래중" ? "활성" : "비활성"}
                  </button>
                </td>
                <td className="border border-slate-200 bg-white px-2 py-1.5 text-right align-middle">
                  <button
                    type="button"
                    onClick={(e) => deleteStore(s, e)}
                    className="rounded-md border border-red-200 bg-red-50 px-2 py-0.5 text-[11px] font-bold text-red-600 hover:bg-red-100"
                    title="거래처 삭제"
                  >
                    X
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      )}
    </div>
  );
}

function TargetByTypePage({ records, setRecords, month }: { records: TargetRecord[]; setRecords: (v: TargetRecord[]) => void; month: string }) {
  const [targetMonth, setTargetMonth] = useState(month);

  const getAmount = (storeType: StoreType) =>
    records
      .filter((r) => r.month === targetMonth && (r.storeType === storeType || (!r.storeType && storeType === "비매장")))
      .reduce((a, b) => a + b.amount, 0);

  const storeAmount = getAmount("매장");
  const nonStoreAmount = getAmount("비매장");

  function updateAmount(storeType: StoreType, amount: number) {
    const next = { storeType, month: targetMonth, amount };
    setRecords([
      ...records.filter((r) => !(r.month === targetMonth && (r.storeType === storeType || (!r.storeType && storeType === "비매장")))),
      next,
    ]);
  }

  async function upload(file: File | null) {
    if (!file) return;
    const fileRows = await readFileRows(file);
    const parsed = fileRows
      .map((r) => {
        const rawType = norm(r["매장구분"] ?? r["구분"] ?? r["타입"] ?? r["storeType"]);
        const storeType: StoreType = rawType.includes("매장") && !rawType.includes("비매장") ? "매장" : "비매장";
        return {
          storeType,
          month: monthText(r["기준월"] ?? r["월"] ?? r["년월"] ?? targetMonth),
          amount: num(r["Target"] ?? r["TARGET"] ?? r["금액"] ?? r["목표"] ?? r["목표매출"]),
        };
      })
      .filter((r) => r.month && r.amount > 0);

    setRecords([
      ...records.filter((r) => !parsed.some((p) => p.month === r.month && p.storeType === r.storeType)),
      ...parsed,
    ]);
    alert(`Target ${parsed.length}건을 반영했습니다.`);
  }

  const excelRows = [
    { 기준월: targetMonth, 구분: "매장", Target: storeAmount },
    { 기준월: targetMonth, 구분: "비매장", Target: nonStoreAmount },
    { 기준월: targetMonth, 구분: "총 Target", Target: storeAmount + nonStoreAmount },
  ];

  return (
    <div className="rounded-xl border border-slate-200/70 bg-white/80 p-4 shadow-sm backdrop-blur">
      <div className="mb-3 flex flex-col gap-2 lg:flex-row lg:items-center lg:justify-between">
        <div>
          <h2 className="text-lg font-bold">Target 관리</h2>
          <p className="mt-1 text-sm text-slate-500">Target은 거래처별이 아니라 매장 / 비매장 기준으로만 관리합니다.</p>
        </div>
        <div className="flex flex-wrap gap-2">
          <button onClick={() => exportExcel(excelRows, `Target관리_${targetMonth}`)} className="rounded-xl bg-emerald-600 px-4 py-2 text-sm font-semibold text-white shadow-sm hover:bg-emerald-700">
            엑셀 다운로드
          </button>
          <label className="cursor-pointer rounded-lg bg-green-600 px-3 py-1.5 text-xs font-semibold text-white">
            엑셀 업로드
            <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={(e) => upload(e.target.files?.[0] || null)} />
          </label>
        </div>
      </div>

      <div className="mb-4 flex flex-wrap gap-2">
        <input type="month" value={targetMonth} onChange={(e) => setTargetMonth(e.target.value)} className="rounded-xl border px-3 py-2" />
      </div>

      <div className="grid grid-cols-1 gap-4 lg:grid-cols-3">
        <div className="rounded-2xl border border-blue-100 bg-blue-50/70 p-4">
          <p className="text-sm font-semibold text-slate-600">매장 Target</p>
          <input value={storeAmount ? won(storeAmount) : ""} onChange={(e) => updateAmount("매장", num(e.target.value))} className="mt-3 w-full rounded-xl border border-slate-300 bg-white px-3 py-3 text-right text-xl font-bold text-slate-900 outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100" placeholder="0" />
        </div>
        <div className="rounded-2xl border border-amber-100 bg-amber-50/70 p-4">
          <p className="text-sm font-semibold text-slate-600">비매장 Target</p>
          <input value={nonStoreAmount ? won(nonStoreAmount) : ""} onChange={(e) => updateAmount("비매장", num(e.target.value))} className="mt-3 w-full rounded-xl border border-slate-300 bg-white px-3 py-3 text-right text-xl font-bold text-slate-900 outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100" placeholder="0" />
        </div>
        <div className="rounded-2xl border border-slate-200 bg-slate-50 p-4">
          <p className="text-sm font-semibold text-slate-600">총 Target</p>
          <p className="mt-3 break-all text-right text-2xl font-bold text-slate-900">{won(storeAmount + nonStoreAmount)}</p>
        </div>
      </div>
    </div>
  );
}

function TargetOrEstPage({ title, records, setRecords, stores, month }: { title: string; records: TargetRecord[] | EstRecord[]; setRecords: (v: any[]) => void; stores: Store[]; month: string }) {
  const [targetMonth, setTargetMonth] = useState(month);
  const [search, setSearch] = useState("");
  const [searchDraft, setSearchDraft] = useState("");
  const isEstPage = title.includes("EST");
  const stMap = storeMap(stores);

  const rows = stores.filter((s) => s.status === "거래중").filter((s) => `${s.code} ${s.name} ${s.manager} ${s.channel}`.toLowerCase().includes(search.toLowerCase())).map((s) => {
    const existing = records.find((r) => r.storeCode === s.code && r.month === targetMonth);
    return { store: s, amount: existing?.amount || 0 };
  });

  const estSummary = records
    .filter((r) => r.month === targetMonth)
    .reduce((acc, r) => {
      const type = stMap.get(r.storeCode || "")?.storeType === "매장" ? "매장" : "비매장";
      if (type === "매장") acc.store += r.amount || 0;
      else acc.nonStore += r.amount || 0;
      return acc;
    }, { store: 0, nonStore: 0 });

  function resetEst() {
    if (!confirm(`${targetMonth} EST 데이터를 초기화할까요?`)) return;
    setRecords(records.filter((r) => r.month !== targetMonth));
    alert(`${targetMonth} EST 데이터가 초기화되었습니다.`);
  }

  function updateAmount(store: Store, amount: number) {
    const next = { storeCode: store.code, storeName: store.name, month: targetMonth, amount };
    setRecords([...records.filter((r) => !(r.storeCode === store.code && r.month === targetMonth)), next]);
  }

  async function upload(file: File | null) {
    if (!file) return;
    const fileRows = await readFileRows(file);
    const parsed = fileRows.map((r) => ({
      storeCode: norm(r["거래처코드"] ?? r["매장코드"]),
      storeName: norm(r["거래처명"] ?? r["매장명"]),
      month: monthText(r["기준월"] ?? r["월"] ?? r["년월"] ?? targetMonth),
      amount: num(r["당월 EST"] ?? r["EST"] ?? r["금액"] ?? r["Target"] ?? r["TARGET"] ?? r["목표"] ?? r["목표매출"]),
    })).filter((r) => r.storeCode && r.month && r.amount > 0);

    setRecords([...records.filter((r) => !parsed.some((p) => p.storeCode === r.storeCode && p.month === r.month)), ...parsed]);
    alert(`${title} ${parsed.length}건을 반영했습니다.`);
  }

  return (
    <div className="rounded-xl border border-slate-200/70 bg-white/80 p-4 shadow-sm backdrop-blur">
      <div className="mb-3 flex flex-col gap-2 lg:flex-row lg:items-center lg:justify-between">
        <div>
          <h2 className="text-lg font-bold">{title}</h2>
          <p className="mt-1 text-sm text-slate-500">거래처별 {title.replace(" 관리", "")} 금액을 관리합니다.</p>
        </div>
        <div className="flex flex-wrap gap-2">
          <label className="cursor-pointer rounded-lg bg-green-600 px-3 py-1.5 text-xs font-semibold text-white">
            엑셀 업로드
            <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={(e) => upload(e.target.files?.[0] || null)} />
          </label>
          {isEstPage && (
            <button type="button" onClick={resetEst} className="rounded-lg border border-red-200 bg-red-50 px-3 py-1.5 text-xs font-semibold text-red-600 hover:bg-red-100">
              EST 리셋
            </button>
          )}
        </div>
      </div>

      {isEstPage && (
        <div className="mb-4 grid grid-cols-1 gap-3 lg:grid-cols-3">
          <div className="rounded-2xl border border-orange-200 bg-orange-50 p-4 shadow-sm">
            <p className="text-sm font-semibold text-slate-600">매장 EST</p>
            <p className="mt-2 break-all text-right text-2xl font-bold text-slate-900">{won(estSummary.store)}</p>
          </div>
          <div className="rounded-2xl border border-orange-200 bg-orange-50 p-4 shadow-sm">
            <p className="text-sm font-semibold text-slate-600">비매장 EST</p>
            <p className="mt-2 break-all text-right text-2xl font-bold text-slate-900">{won(estSummary.nonStore)}</p>
          </div>
          <div className="rounded-2xl border border-orange-300 bg-orange-100 p-4 shadow-sm">
            <p className="text-sm font-semibold text-slate-600">총 EST</p>
            <p className="mt-2 break-all text-right text-2xl font-bold text-slate-900">{won(estSummary.store + estSummary.nonStore)}</p>
          </div>
        </div>
      )}

      <div className="mb-4 flex flex-wrap gap-2">
        <input type="month" value={targetMonth} onChange={(e) => setTargetMonth(e.target.value)} className="rounded-xl border px-3 py-2" />
        <input
          value={searchDraft}
          onChange={(e) => {
            const value = e.target.value;
            setSearchDraft(value);
            if (!value.trim()) setSearch("");
          }}
          onKeyDown={(e) => {
            if (e.key === "Enter") setSearch(searchDraft.trim());
          }}
          placeholder="거래처/담당자/채널 검색"
          className="w-[280px] rounded-xl border px-3 py-2"
        />
      </div>

      <div className="max-h-[62vh] overflow-auto">
        <table className="w-full min-w-[820px] border text-sm">
          <thead>
            <tr className="bg-slate-100">
              <Th>거래처코드</Th>
              <Th>거래처명</Th>
              <Th>채널</Th>
              <Th>담당자</Th>
              <Th right>금액</Th>
            </tr>
          </thead>
          <tbody>
            {rows.map(({ store, amount }) => (
              <tr key={store.code}>
                <Td>{store.code}</Td>
                <Td bold>{store.name}</Td>
                <Td>{store.channel}</Td>
                <Td>{store.manager || "-"}</Td>
                <Td right>
                  <input
                    value={amount ? won(amount) : ""}
                    onChange={(e) => updateAmount(store, num(e.target.value))}
                    className="w-[160px] rounded border px-2 py-1 text-right"
                    placeholder="0"
                  />
                </Td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}


function buildAutoClosedStoresFromPrevYear(parsed: SalesRecord[], stores: Store[]) {
  const existingByCode = new Map(stores.map((s) => [norm(s.code), s]));
  const currentActiveByName = new Set(
    stores
      .filter((s) => s.status === "거래중")
      .map((s) => normalizeStoreNameKey(s.name))
      .filter(Boolean)
  );
  const map = new Map<string, Store>();

  parsed.forEach((r) => {
    const codeKey = norm(r.storeCode);
    const nameKey = normalizeStoreNameKey(r.storeName);
    if (!codeKey || existingByCode.has(codeKey) || (nameKey && currentActiveByName.has(nameKey))) return;

    map.set(codeKey, {
      code: r.storeCode,
      name: r.storeName || r.storeCode,
      channel: r.channel || "미지정",
      manager: (r.manager || "") as Manager,
      storeType: r.storeType || "비매장",
      brand: r.brand || "미지정",
      status: "거래종료",
    });
  });

  return Array.from(map.values());
}

function UploadPage({ stores, setStores, sales, setSales, month, date, timeConfigs, setTimeConfigs }: { stores: Store[]; setStores: (v: Store[]) => void; sales: SalesRecord[]; setSales: React.Dispatch<React.SetStateAction<SalesRecord[]>>; month: string; date: string; timeConfigs: TimeConfig[]; setTimeConfigs: (v: TimeConfig[]) => void }) {
  const [holidayText, setHolidayText] = useState("");
  const [deleteDate, setDeleteDate] = useState(today());

  async function uploadSales(file: File | null, period: PeriodType) {
    if (!file) return;
    const rows = await readFileRows(file);
    const parsed = rows.map((r, index) => {
      const saleDate = dateText(r["일자 No."] ?? r["일자"] ?? r["매출일"] ?? r["판매일"] ?? r["기준일"] ?? deleteDate);
      const storeCode = norm(r["거래처 코드"] ?? r["거래처코드"] ?? r["매장코드"]);
      const storeName = norm(r["거래처명"] ?? r["매장명"]);
      const itemCode = norm(r["품목 코드"] ?? r["품목코드"] ?? r["상품코드"]);
      const itemName = norm(r["품목명[규격]"] ?? r["품목명"] ?? r["상품명"]);
      const quantity = num(r["판매 수량"] ?? r["수량"]);
      const rawSalesAmount = r["판매 금액"] ?? r["판매금액"] ?? r["매출금액"] ?? r["당월 매출"] ?? r["매출"];
      const rawCostAmount = r["원가 금액"] ?? r["원가금액"] ?? r["원가"];
      const rawProfitAmount = r["이익 금액"] ?? r["이익금액"] ?? r["매출 이익"] ?? r["매출이익"] ?? r["매출총이익"] ?? r["마진액"];
      const hasSalesAmount = norm(rawSalesAmount) !== "";
      const hasCostAmount = norm(rawCostAmount) !== "";
      const hasProfitAmount = norm(rawProfitAmount) !== "";
      const salesAmount = num(rawSalesAmount);
      const costAmount = num(rawCostAmount);
      const uploadedProfitAmount = num(rawProfitAmount);
      const profitAmount = hasProfitAmount
        ? uploadedProfitAmount
        : hasSalesAmount || hasCostAmount
        ? salesAmount - costAmount
        : 0;
      const uploadedProfitRate = profitRateValue(
        r["이익율"] ??
        r["이익률"] ??
        r["매출 이익율"] ??
        r["매출 이익률"] ??
        r["매출이익율"] ??
        r["매출이익률"] ??
        r["마진율"]
      );
      const mapping = storeMap(stores).get(storeCode);

      return makeSale(
        period,
        month,
        saleDate,
        storeCode,
        storeName,
        itemCode || `ITEM-${index}`,
        itemName,
        quantity,
        salesAmount,
        costAmount,
        profitAmount,
        stores,
        uploadedProfitRate
      );
    }).filter((r) =>
      r.saleDate &&
      r.storeCode &&
      (r.salesAmount !== 0 || r.costAmount !== 0 || r.profitAmount !== 0 || r.quantity !== 0)
    );

    const missingStores = period === "prevYear"
      ? buildAutoClosedStoresFromPrevYear(parsed, stores)
      : parsed
        .filter((r) => !storeMap(stores).has(r.storeCode))
        .map((r) => ({
          code: r.storeCode,
          name: r.storeName || r.storeCode,
          channel: "매장" as Channel,
          manager: "" as Manager,
          storeType: "매장" as StoreType,
          brand: r.brand || "미지정",
          status: "거래중" as const,
        }));

    if (missingStores.length) {
      const map = new Map(stores.map((s) => [s.code, s]));
      missingStores.forEach((s) => map.set(s.code, s));
      setStores(Array.from(map.values()));
    }

    if (parsed.length === 0) {
      alert("업로드 파일에서 반영할 매출/원가/이익 행을 찾지 못했습니다. 기존 데이터는 삭제하지 않았습니다.");
      return;
    }

    const uploadedDates = Array.from(new Set(parsed.map((r) => r.saleDate).filter(Boolean)));

    // 업로드 직후 금액이 들어갔다가 사라지는 문제를 막기 위해
    // 화면에 잡혀 있던 오래된 sales 값이 아니라, setSales가 받는 최신 prev 값을 기준으로 병합합니다.
    setSales((prevSales) => {
      const nextSales = period === "current"
        ? prevSales.filter((s) => !(s.period === "current" && uploadedDates.includes(s.saleDate)))
        : prevSales.filter((s) => !(s.period === period && s.refMonth === month));

      return [...nextSales, ...parsed];
    });

    const closedMessage = period === "prevYear" && missingStores.length ? `\n당월 기준에 없는 전년동월 거래처 ${missingStores.length}건은 거래종료로 자동 생성했습니다.` : "";
    alert(`${period === "current" ? "당월" : period === "prevMonth" ? "전월" : "전년동월"} 매출 ${parsed.length}건을 반영했습니다.\n반영 날짜: ${uploadedDates.join(", ")}${closedMessage}`);
  }

  function saveHolidays() {
    const holidays: string[] = Array.from(new Set<string>(holidayText.split(/\n|,|\s+/).map(dateText).filter((d): d is string => Boolean(d && d.startsWith(month))))).sort();
    setTimeConfigs([...timeConfigs.filter((c) => c.month !== month), { month, holidays }]);
    alert(`${month} TIME GONE 공휴일 ${holidays.length}건을 저장했습니다.`);
  }

  function deleteCurrentDate() {
    if (!confirm(`${deleteDate} 당월 매출 데이터를 삭제할까요?`)) return;
    setSales(sales.filter((s) => !(s.period === "current" && s.saleDate === deleteDate)));
  }

  return (
    <div className="space-y-4">
      <div className="rounded-2xl border border-slate-200/70 bg-white/80 p-5 shadow-sm backdrop-blur">
        <h2 className="mb-3 text-lg font-bold">매출 업로드</h2>
        <div className="grid grid-cols-1 gap-4 lg:grid-cols-3">
          <UploadBox title="당월 매출 업로드" description="같은 날짜 파일을 다시 올리면 해당 날짜 기존 당월 매출을 삭제하고 새 파일로 업데이트합니다." onUpload={(file) => uploadSales(file, "current")} />
          <UploadBox title="전월 매출 업로드" description="매출비교와 매출현황의 전월 매출 기준으로 사용합니다. 같은 기준월 자료는 새 파일로 교체됩니다." onUpload={(file) => uploadSales(file, "prevMonth")} />
          <UploadBox title="전년동월 매출 업로드" description="매출비교와 매출현황의 전년동월 매출 기준으로 사용합니다. 같은 기준월 자료는 새 파일로 교체됩니다." onUpload={(file) => uploadSales(file, "prevYear")} />
        </div>
      </div>

      <ProfitValidationPanel sales={sales} month={month} date={date} />

      <div className="rounded-2xl border border-slate-200/70 bg-white/80 p-5 shadow-sm backdrop-blur">
        <h2 className="mb-3 text-lg font-bold">당월 특정 날짜 삭제</h2>
        <input type="date" value={deleteDate} onChange={(e) => setDeleteDate(e.target.value)} className="mr-2 rounded-xl border px-3 py-2" />
        <button onClick={deleteCurrentDate} className="rounded-xl bg-red-600 px-4 py-2 text-white">해당일 당월 매출 삭제</button>
      </div>

      <div className="rounded-2xl border border-slate-200/70 bg-white/80 p-5 shadow-sm backdrop-blur">
        <h2 className="mb-3 text-lg font-bold">TIME GONE 공휴일 설정</h2>
        <p className="mb-3 text-sm text-slate-500">월~금 일반일 1일, 월~금 공휴일 0.5일, 토요일 0.5일, 일요일 0일 기준입니다.</p>
        <textarea value={holidayText} onChange={(e) => setHolidayText(e.target.value)} placeholder={`${month}-06\n${month}-15`} className="h-28 w-full rounded-xl border px-3 py-2" />
        <button onClick={saveHolidays} className="mt-3 rounded-xl bg-blue-600 px-4 py-2 text-white">공휴일 저장</button>
      </div>
    </div>
  );
}


function ProfitValidationPanel({ sales, month, date }: { sales: SalesRecord[]; month: string; date: string }) {
  const monthRows = sales.filter((s) => s.period === "current" && inRange(s.saleDate, monthStart(month), monthEnd(month)));
  const toDateRows = sales.filter((s) => s.period === "current" && inRange(s.saleDate, monthStart(month), date));
  const monthProfit = sum(monthRows, "profitAmount");
  const toDateProfit = sum(toDateRows, "profitAmount");
  const monthSales = sum(monthRows, "salesAmount");
  const toDateSales = sum(toDateRows, "salesAmount");
  const monthRate = weightedProfitRate(monthRows);
  const toDateRate = weightedProfitRate(toDateRows);

  const dailyRows: { date: string; salesAmount: number; profitAmount: number; profitRate: number }[] = [];
  for (let d = monthStart(month), guard = 0; d <= monthEnd(month) && guard < 40; d = addDays(d, 1), guard += 1) {
    const rows = monthRows.filter((s) => s.saleDate === d);
    dailyRows.push({
      date: d,
      salesAmount: sum(rows, "salesAmount"),
      profitAmount: sum(rows, "profitAmount"),
      profitRate: weightedProfitRate(rows),
    });
  }

  const excelRows = dailyRows.map((r) => ({
    날짜: r.date,
    매출금액: r.salesAmount,
    이익금액: r.profitAmount,
    이익률: pct(r.profitRate),
  }));

  return (
    <div className="rounded-2xl border border-slate-200/70 bg-white/80 p-5 shadow-sm backdrop-blur">
      <div className="mb-3 flex flex-col gap-2 lg:flex-row lg:items-center lg:justify-between">
        <div>
          <h2 className="text-lg font-bold">이익금액 검증</h2>
          <p className="mt-1 text-sm text-slate-500">당월 전체와 기준일({date})까지의 이익금액을 나누어 확인합니다.</p>
        </div>
        <button
          type="button"
          onClick={() => exportExcel(excelRows, `이익금액_날짜별_검증_${month}`)}
          className="rounded-xl bg-emerald-600 px-4 py-2 text-sm font-semibold text-white shadow-sm hover:bg-emerald-700"
        >
          날짜별 검증표 다운로드
        </button>
      </div>

      <div className="mb-4 grid grid-cols-1 gap-3 md:grid-cols-2">
        <div className="rounded-xl border border-green-200 bg-green-50 p-4">
          <p className="text-sm font-semibold text-slate-600">당월 이익금액</p>
          <p className="mt-2 text-2xl font-bold text-slate-900">{won(monthProfit)}원</p>
          <p className="mt-1 text-xs text-slate-500">당월 매출 {won(monthSales)}원 · 이익률 {pct(monthRate)}</p>
        </div>
        <div className="rounded-xl border border-amber-200 bg-amber-50 p-4">
          <p className="text-sm font-semibold text-slate-600">당일까지의 이익</p>
          <p className="mt-2 text-2xl font-bold text-slate-900">{won(toDateProfit)}원</p>
          <p className="mt-1 text-xs text-slate-500">당일까지 매출 {won(toDateSales)}원 · 이익률 {pct(toDateRate)}</p>
        </div>
      </div>

      <div className="max-h-[360px] overflow-auto">
        <table className="w-full min-w-[640px] border border-slate-200 text-xs">
          <thead>
            <tr className="bg-slate-100">
              <th className="sticky top-0 border bg-slate-100 px-3 py-2 text-left font-bold">날짜</th>
              <th className="sticky top-0 border bg-slate-100 px-3 py-2 text-right font-bold">매출금액</th>
              <th className="sticky top-0 border bg-slate-100 px-3 py-2 text-right font-bold">이익금액</th>
              <th className="sticky top-0 border bg-slate-100 px-3 py-2 text-right font-bold">이익률</th>
            </tr>
          </thead>
          <tbody>
            {dailyRows.map((r) => (
              <tr key={r.date} className={r.date <= date ? "bg-white" : "bg-slate-50 text-slate-500"}>
                <td className="border px-3 py-2 font-semibold">{r.date}</td>
                <td className="border px-3 py-2 text-right font-semibold text-slate-900">{won(r.salesAmount)}</td>
                <td className="border px-3 py-2 text-right font-semibold text-slate-900">{won(r.profitAmount)}</td>
                <td className="border px-3 py-2 text-right">{pct(r.profitRate)}</td>
              </tr>
            ))}
          </tbody>
          <tfoot>
            <tr className="bg-slate-100 font-bold">
              <td className="border px-3 py-2">합계</td>
              <td className="border px-3 py-2 text-right">{won(monthSales)}</td>
              <td className="border px-3 py-2 text-right">{won(monthProfit)}</td>
              <td className="border px-3 py-2 text-right">{pct(monthRate)}</td>
            </tr>
          </tfoot>
        </table>
      </div>
    </div>
  );
}

function UploadBox({ title, description, onUpload }: { title: string; description: string; onUpload: (file: File | null) => void }) {
  return (
    <div className="rounded-2xl border border-slate-200 bg-slate-50/70 p-4">
      <h3 className="font-bold text-slate-900">{title}</h3>
      <p className="mt-2 min-h-[44px] text-sm text-slate-500">{description}</p>
      <label className="mt-4 inline-block cursor-pointer rounded-xl bg-blue-600 px-4 py-2 text-sm font-semibold text-white">
        엑셀 업로드
        <input
          type="file"
          accept=".xlsx,.xls,.csv"
          className="hidden"
          onChange={(e) => {
            const file = e.target.files?.[0] || null;
            onUpload(file);
            e.currentTarget.value = "";
          }}
        />
      </label>
    </div>
  );
}

function Th({ children, right = false, tone = "default" }: { children: React.ReactNode; right?: boolean; tone?: "default" | "mint" | "blue" | "yellow" | "gray" | "purple" | "green" }) {
  const toneClass =
    tone === "mint" ? "border-emerald-200 bg-emerald-50 text-emerald-800" :
    tone === "blue" ? "border-blue-200 bg-blue-50 text-blue-800" :
    tone === "yellow" ? "border-amber-200 bg-amber-50 text-amber-800" :
    tone === "gray" ? "border-slate-300 bg-slate-100 text-slate-800" :
    tone === "purple" ? "border-violet-200 bg-violet-50 text-violet-800" :
    tone === "green" ? "border-green-200 bg-green-50 text-green-800" :
    "border-slate-200 bg-slate-100 text-slate-800";

  return <th className={`sticky top-0 z-20 border p-2 text-xs font-bold leading-tight whitespace-normal break-keep ${toneClass} ${right ? "text-right" : "text-left"}`}>{children}</th>;
}

function Td({ children, right = false, bold = false, color = "" }: { children: React.ReactNode; right?: boolean; bold?: boolean; color?: string }) {
  return <td className={`border border-slate-200 bg-white p-2 ${right ? "text-right" : "text-left"} ${bold ? "font-semibold" : ""} ${color}`}>{children}</td>;
}
