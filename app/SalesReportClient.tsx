"use client";

import { Fragment, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

type Channel = string;
type Manager = string;
type StoreType = string;
type PeriodType = "current" | "prevMonth" | "prevYear";
type SalesView = "거래처별" | "브랜드별" | "담당자별" | "채널별";
type MonthStartTab = "거래처 리스트" | "기준정보" | "업로드 관리" | "이익금액 검증표";
type DrillPeriod = "prevYear" | "prevMonth" | "current" | "currentFullMonth";
type SalesStatusSortKey =
  | "label"
  | "prevYearSales"
  | "prevYearRate"
  | "prevYearTimeGoneGap"
  | "prevMonthSales"
  | "prevMonthRate"
  | "prevMonthTimeGoneGap"
  | "currentSales"
  | "fullMonthSales"
  | "timeGone"
  | "timeGoneGap"
  | "est"
  | "estRate"
  | "profitAmount"
  | "profitRate"
  | "lastOrderDate";
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

type SalesUploadRequest = {
  period: PeriodType;
  refMonth: string;
  fileName: string;
  uploadedDates: string[];
  rows: SalesRecord[];
};

type SalesStorageActions = {
  replaceUpload: (request: SalesUploadRequest) => Promise<{ mode: "v3" | "legacy" }>;
  deleteCurrentDate: (refMonth: string, saleDate: string) => Promise<{ mode: "v3" | "legacy" }>;
  refresh: () => Promise<void>;
  storageMode: "checking" | "v3" | "legacy";
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

type ItemCostHistory = {
  id: string;
  changedAt: string;
  effectiveDate: string;
  previousCost: number;
  newCost: number;
  memo?: string;
};

type ItemMasterRecord = {
  itemCode: string;
  itemName: string;
  category: string;
  active?: boolean;
  memo?: string;
};

type ItemCostRecord = {
  itemCode: string;
  itemName: string;
  currentCost: number;
  nextCost?: number;
  effectiveDate?: string;
  memo?: string;
  history: ItemCostHistory[];
};

type TimeConfig = {
  month: string;
  holidays: string[];
};

const CHANNELS: Channel[] = [
  "도매",
  "체인",
  "체인물류",
  "식자재마트",
  "제조",
  "권역배송",
  "온라인",
  "매장",
  "비매장",
  "기업",
  "매입",
  "본사",
];
const MANAGERS: Manager[] = ["SY", "KT", "SW", "NH", "Bomi", "BM", "bomi"];
const SALES_VIEWS: SalesView[] = ["거래처별", "브랜드별", "담당자별"];
const MONTH_TABS: MonthStartTab[] = [
  "거래처 리스트",
  "기준정보",
  "업로드 관리",
  "이익금액 검증표",
];

const initialStores: Store[] = [
  {
    code: "385-81-04167",
    name: "주식회사 메이트레이더스",
    channel: "도매",
    manager: "SY",
    storeType: "비매장",
    brand: "이마트 (수출)",
    status: "거래중",
  },
  {
    code: "1298688941-1",
    name: "(주)에스피씨 지에프에스_왓더버거",
    channel: "체인",
    manager: "KT",
    storeType: "비매장",
    brand: "SPC(왓더버거)",
    status: "거래중",
  },
  {
    code: "6018700459",
    name: "명랑시대외식청년창업협동조합",
    channel: "체인",
    manager: "SY",
    storeType: "비매장",
    brand: "명랑시대",
    status: "거래중",
  },
  {
    code: "4678601074-3",
    name: "세이웰_왁버거(제때)",
    channel: "체인물류",
    manager: "SY",
    storeType: "비매장",
    brand: "세이웰",
    status: "거래중",
  },
  {
    code: "417-81-24010",
    name: "주식회사 해창수산",
    channel: "도매",
    manager: "KT",
    storeType: "비매장",
    brand: "주식회사 해창수산",
    status: "거래중",
  },
  {
    code: "237-88-02985",
    name: "주식회사 하라에프에스",
    channel: "도매",
    manager: "SY",
    storeType: "비매장",
    brand: "주식회사 하라에프에스",
    status: "거래중",
  },
  {
    code: "1368702986",
    name: "주식회사 쓰담",
    channel: "제조",
    manager: "SY",
    storeType: "비매장",
    brand: "주식회사 쓰담",
    status: "거래중",
  },
  {
    code: "119-86-48020",
    name: "더한솔씨앤에스(주)",
    channel: "체인",
    manager: "SY",
    storeType: "비매장",
    brand: "더한솔씨앤에스(주)",
    status: "거래중",
  },
  {
    code: "999-1",
    name: "[온라인] 스마트스토어 고객",
    channel: "온라인",
    manager: "Bomi",
    storeType: "비매장",
    brand: "온라인",
    status: "거래중",
  },
  {
    code: "447-81-01963",
    name: "주식회사 메이코더스",
    channel: "도매",
    manager: "SY",
    storeType: "비매장",
    brand: "이마트 (수출)",
    status: "거래중",
  },
  {
    code: "4678601074-2",
    name: "세이웰_번패티번(CJFW)",
    channel: "체인물류",
    manager: "SY",
    storeType: "비매장",
    brand: "세이웰",
    status: "거래중",
  },
  {
    code: "2598501011",
    name: "(주)푸드엔 물류센터",
    channel: "식자재마트",
    manager: "KT",
    storeType: "비매장",
    brand: "(주)푸드엔 물류센터",
    status: "거래중",
  },
  {
    code: "5038701038",
    name: "주식회사 조인앤조인",
    channel: "제조",
    manager: "SY",
    storeType: "비매장",
    brand: "주식회사 조인앤조인",
    status: "거래중",
  },
  {
    code: "268-88-02001",
    name: "세븐패티버거 을지로본점",
    channel: "체인",
    manager: "SW",
    storeType: "매장",
    brand: "세븐패티버거",
    status: "거래중",
  },
  {
    code: "8078802372",
    name: "주식회사 링커",
    channel: "도매",
    manager: "SY",
    storeType: "비매장",
    brand: "주식회사 링커",
    status: "거래중",
  },
  {
    code: "4638501852",
    name: "온더보더여의도 IFC점 (SRW)",
    channel: "체인",
    manager: "SY",
    storeType: "비매장",
    brand: "온더보더",
    status: "거래중",
  },
  {
    code: "2028535854",
    name: "온더보더 광화문 D타워점 (SRW)",
    channel: "체인",
    manager: "SY",
    storeType: "비매장",
    brand: "온더보더",
    status: "거래중",
  },
  {
    code: "6765800749",
    name: "더블 식스 버거",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "더블 식스 버거",
    status: "거래중",
  },
  {
    code: "8262002067",
    name: "스니커스(snickers)홍대",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "스니커스(snickers)홍대",
    status: "거래중",
  },
  {
    code: "268-88-02002",
    name: "세븐패티버거 압구정로데오점",
    channel: "체인",
    manager: "SW",
    storeType: "매장",
    brand: "세븐패티버거",
    status: "거래중",
  },
  {
    code: "4678601074",
    name: "주식회사 세이웰",
    channel: "체인물류",
    manager: "SY",
    storeType: "비매장",
    brand: "세이웰",
    status: "거래중",
  },
  {
    code: "1208548703",
    name: "온더보더 코엑스 도심공항점 (JRW)",
    channel: "체인",
    manager: "SY",
    storeType: "비매장",
    brand: "온더보더",
    status: "거래중",
  },
  {
    code: "4206400773",
    name: "세븐패티버거 송리단길점",
    channel: "체인",
    manager: "SW",
    storeType: "매장",
    brand: "세븐패티버거",
    status: "거래중",
  },
  {
    code: "7518500503",
    name: "온더보더 스타필드하남점 (JRW)",
    channel: "체인",
    manager: "SY",
    storeType: "비매장",
    brand: "온더보더",
    status: "거래중",
  },
  {
    code: "2131641419",
    name: "세븐패티버거 청량리점",
    channel: "체인",
    manager: "SW",
    storeType: "매장",
    brand: "세븐패티버거",
    status: "거래중",
  },
  {
    code: "5392801821",
    name: "필앤필버거 김포점",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "필앤필버거 김포점",
    status: "거래중",
  },
  {
    code: "6038111772",
    name: "CJFW(주)부산센터",
    channel: "기업",
    manager: "KT",
    storeType: "비매장",
    brand: "CJFW",
    status: "거래중",
  },
  {
    code: "189-81-00700",
    name: "주식회사 굿프랜즈",
    channel: "제조",
    manager: "SY",
    storeType: "비매장",
    brand: "주식회사 굿프랜즈",
    status: "거래중",
  },
  {
    code: "1078540256",
    name: "온더보더 타임스퀘어점 (JRW)",
    channel: "체인",
    manager: "SY",
    storeType: "비매장",
    brand: "온더보더",
    status: "거래중",
  },
  {
    code: "2968500520",
    name: "온더보더 롯데몰김포공항점 (JRW)",
    channel: "체인",
    manager: "SY",
    storeType: "비매장",
    brand: "온더보더",
    status: "거래중",
  },
  {
    code: "6848502439",
    name: "후라이드참잘하는집",
    channel: "체인",
    manager: "SY",
    storeType: "비매장",
    brand: "후라이드참잘하는집",
    status: "거래중",
  },
  {
    code: "2102974658",
    name: "벅벅버거",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "벅벅버거",
    status: "거래중",
  },
  {
    code: "2896500617",
    name: "매드버거",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "매드버거",
    status: "거래중",
  },
  {
    code: "3103801103",
    name: "필앤필버거(PHIL&FILL BURGER)",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "필앤필버거(PHIL&FILL BURGER)",
    status: "거래중",
  },
  {
    code: "1877200148",
    name: "버기즈 탄방점",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "버기즈 탄방점",
    status: "거래중",
  },
  {
    code: "3621002813",
    name: "크레이지버거(가람점)",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "크레이지버거",
    status: "거래중",
  },
  {
    code: "5332301911",
    name: "크레이지버거 밤리단길 일산 밤가시점",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "크레이지버거",
    status: "거래중",
  },
  {
    code: "6646700544",
    name: "버기즈 어은점",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "버기즈 어은점",
    status: "거래중",
  },
  {
    code: "3878102559",
    name: "(주)인맥에프엔씨",
    channel: "체인",
    manager: "SY",
    storeType: "비매장",
    brand: "(주)인맥에프엔씨",
    status: "거래중",
  },
  {
    code: "3363101654",
    name: "디어버거 진주경상대점",
    channel: "권역배송",
    manager: "KT",
    storeType: "매장",
    brand: "디어버거",
    status: "거래중",
  },
  {
    code: "6838503624",
    name: "온더보더 하버마스타 (JRW)",
    channel: "체인",
    manager: "SY",
    storeType: "비매장",
    brand: "온더보더",
    status: "거래중",
  },
  {
    code: "3268501475",
    name: "온더보더 대전테크노중앙로점 (JRW)",
    channel: "체인",
    manager: "SY",
    storeType: "비매장",
    brand: "온더보더",
    status: "거래중",
  },
  {
    code: "4686400794",
    name: "버거바",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "버거바",
    status: "거래중",
  },
  {
    code: "3210404046",
    name: "크레이지버거 수원 인계점",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "크레이지버거",
    status: "거래중",
  },
  {
    code: "5244901144",
    name: "세븐패티버거 영등포점",
    channel: "체인",
    manager: "SW",
    storeType: "매장",
    brand: "세븐패티버거",
    status: "거래중",
  },
  {
    code: "6498802761",
    name: "주식회사 로이플렉스",
    channel: "체인",
    manager: "NH",
    storeType: "비매장",
    brand: "주식회사 로이플렉스",
    status: "거래중",
  },
  {
    code: "120-85-52717",
    name: "제스티살룬 네이버점",
    channel: "체인",
    manager: "SW",
    storeType: "매장",
    brand: "제스티살룬",
    status: "거래중",
  },
  {
    code: "2414100337",
    name: "킴스키친",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "킴스키친",
    status: "거래중",
  },
  {
    code: "1562901851",
    name: "디어버거_사천점",
    channel: "권역배송",
    manager: "KT",
    storeType: "매장",
    brand: "디어버거",
    status: "거래중",
  },
  {
    code: "7200601274",
    name: "오일리버거 교동본점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "오일리버거 교동본점",
    status: "거래중",
  },
  {
    code: "1835300669",
    name: "파이어벨 코엑스점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "파이어벨 코엑스점",
    status: "거래중",
  },
  {
    code: "3360103190",
    name: "디어버거 광양점",
    channel: "권역배송",
    manager: "KT",
    storeType: "매장",
    brand: "디어버거",
    status: "거래중",
  },
  {
    code: "999",
    name: "[온라인] (주)마켓보로",
    channel: "온라인",
    manager: "Bomi",
    storeType: "비매장",
    brand: "온라인",
    status: "거래중",
  },
  {
    code: "8654600647",
    name: "요리지존(파스타왕)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "요리지존(파스타왕)",
    status: "거래중",
  },
  {
    code: "5605100805",
    name: "디얼버거(DEAR BURGER) (디어버거 진주)",
    channel: "권역배송",
    manager: "KT",
    storeType: "매장",
    brand: "디어버거",
    status: "거래중",
  },
  {
    code: "6294001446",
    name: "트레인 버거(TRAIN BURGER)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "트레인 버거(TRAIN BURGER)",
    status: "거래중",
  },
  {
    code: "6713401621",
    name: "매드버거 (궁동점)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "매드버거 (궁동점)",
    status: "거래중",
  },
  {
    code: "5618101202",
    name: "주식회사 쉬즈베이글",
    channel: "체인",
    manager: "SY",
    storeType: "비매장",
    brand: "주식회사 쉬즈베이글",
    status: "거래중",
  },
  {
    code: "1288639317",
    name: "㈜ 엔젤푸드",
    channel: "도매",
    manager: "NH",
    storeType: "비매장",
    brand: "㈜ 엔젤푸드",
    status: "거래중",
  },
  {
    code: "2201006816",
    name: "제레미버거_양재본점",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "제레미버거_양재본점",
    status: "거래중",
  },
  {
    code: "5523901096",
    name: "뉴욕버거 김해공항점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "뉴욕버거 김해공항점",
    status: "거래중",
  },
  {
    code: "1020288010",
    name: "행루즈버거",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "행루즈버거",
    status: "거래중",
  },
  {
    code: "6246100786",
    name: "1986 시부야버거",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "1986 시부야버거",
    status: "거래중",
  },
  {
    code: "4392901722",
    name: "버기즈 관평점",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "버기즈 관평점",
    status: "거래중",
  },
  {
    code: "4570603409",
    name: "버거고",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "버거고",
    status: "거래중",
  },
  {
    code: "4841502234",
    name: "버거비앤비(광교점)",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "버거비앤비(광교점)",
    status: "거래중",
  },
  {
    code: "2081721529",
    name: "버거타임",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "버거타임",
    status: "거래중",
  },
  {
    code: "6031594669",
    name: "홀리버거(HOLY BURGER)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "홀리버거(HOLY BURGER)",
    status: "거래중",
  },
  {
    code: "5500203302",
    name: "로우로우 버거샵(RAW RAW Burgershop)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "로우로우 버거샵(RAW RAW Burgershop)",
    status: "거래중",
  },
  {
    code: "6342801091",
    name: "피자팬팬",
    channel: "체인",
    manager: "NH",
    storeType: "매장",
    brand: "피자팬팬",
    status: "거래중",
  },
  {
    code: "1787200426",
    name: "버거비앤비(BURGER BNB)",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "버거비앤비(BURGER BNB)",
    status: "거래중",
  },
  {
    code: "6568102756",
    name: "(주)현대그린푸드",
    channel: "제조",
    manager: "SY",
    storeType: "비매장",
    brand: "(주)현대그린푸드",
    status: "거래중",
  },
  {
    code: "3558503317",
    name: "온더보더 마곡 코엑스점 (JRW)",
    channel: "체인",
    manager: "SY",
    storeType: "비매장",
    brand: "온더보더",
    status: "거래중",
  },
  {
    code: "2068694885",
    name: "버거비",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "버거비",
    status: "거래중",
  },
  {
    code: "2738502132",
    name: "온더보더 에버랜드점 (JRW)",
    channel: "체인",
    manager: "SY",
    storeType: "비매장",
    brand: "온더보더",
    status: "거래중",
  },
  {
    code: "6182492568",
    name: "벅벅버거(이태원점)",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "벅벅버거",
    status: "거래중",
  },
  {
    code: "7958801686",
    name: "(주)테이스테이 (Teistay)_홍대점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "(주)테이스테이 (Teistay)_홍대점",
    status: "거래중",
  },
  {
    code: "2880703192",
    name: "디어버거 도계점",
    channel: "권역배송",
    manager: "KT",
    storeType: "매장",
    brand: "디어버거",
    status: "거래중",
  },
  {
    code: "4938503263",
    name: "온더보더 롯데프리미엄아울렛 동부산점 (SRW)",
    channel: "체인",
    manager: "SY",
    storeType: "비매장",
    brand: "온더보더",
    status: "거래중",
  },
  {
    code: "8436400277",
    name: "제레미버거_군자점",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "제레미버거_군자점",
    status: "거래중",
  },
  {
    code: "5366600349",
    name: "제레미버거_선유도점",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "제레미버거_선유도점",
    status: "거래중",
  },
  {
    code: "2036403175",
    name: "백소정 서창점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "백소정",
    status: "거래중",
  },
  {
    code: "5101766564",
    name: "버거쑈",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "버거쑈",
    status: "거래중",
  },
  {
    code: "6331202228",
    name: "백소정 인하대후문점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "백소정",
    status: "거래중",
  },
  {
    code: "7283701350",
    name: "백소정 주안역점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "백소정",
    status: "거래중",
  },
  {
    code: "8641702604",
    name: "신촌버거",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "신촌버거",
    status: "거래중",
  },
  {
    code: "3142783173",
    name: "폼프리츠 카이스트점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "폼프리츠 카이스트점",
    status: "거래중",
  },
  {
    code: "1858100902",
    name: "주식회사 스모크하우스512",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "주식회사 스모크하우스512",
    status: "거래중",
  },
  {
    code: "2363800785",
    name: "써니사이드 다이닝",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "써니사이드 다이닝",
    status: "거래중",
  },
  {
    code: "7488602763",
    name: "주식회사 제스티살룬",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "주식회사 제스티살룬",
    status: "거래중",
  },
  {
    code: "3951202156",
    name: "백소정 서울대입구역점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "백소정",
    status: "거래중",
  },
  {
    code: "7958801002",
    name: "(주)테이스테이_마곡점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "(주)테이스테이_마곡점",
    status: "거래중",
  },
  {
    code: "5410200975",
    name: "소울버킷(SOUL BUCKET)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "소울버킷(SOUL BUCKET)",
    status: "거래중",
  },
  {
    code: "3601502660",
    name: "세븐패티버거 낙성대점",
    channel: "체인",
    manager: "SW",
    storeType: "매장",
    brand: "세븐패티버거",
    status: "거래중",
  },
  {
    code: "7333701314",
    name: "버거치즈스마일 원곡점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "버거치즈스마일 원곡점",
    status: "거래중",
  },
  {
    code: "6384700741",
    name: "보스턴 김포 장기점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "보스턴 김포 장기점",
    status: "거래중",
  },
  {
    code: "5182601650",
    name: "백소정 옥수점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "백소정",
    status: "거래중",
  },
  {
    code: "2346900196",
    name: "버거401",
    channel: "권역배송",
    manager: "KT",
    storeType: "매장",
    brand: "버거401",
    status: "거래중",
  },
  {
    code: "999-2",
    name: "[온라인] 쿠팡 고객",
    channel: "온라인",
    manager: "bomi",
    storeType: "비매장",
    brand: "온라인",
    status: "거래중",
  },
  {
    code: "6990103779",
    name: "떰즈업 (THUMBS UP)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "떰즈업 (THUMBS UP)",
    status: "거래중",
  },
  {
    code: "6848500002",
    name: "후참잘 (케이닥)",
    channel: "체인",
    manager: "SY",
    storeType: "비매장",
    brand: "후라이드참잘하는집",
    status: "거래중",
  },
  {
    code: "2252901370",
    name: "백소정 과천점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "백소정",
    status: "거래중",
  },
  {
    code: "8330202630",
    name: "백소정 구월로데오점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "백소정",
    status: "거래중",
  },
  {
    code: "2050815666",
    name: "도그즈인번즈(Dogs Buns)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "도그즈인번즈(Dogs Buns)",
    status: "거래중",
  },
  {
    code: "4383401494",
    name: "벅벅버거 (신당점)",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "벅벅버거",
    status: "거래중",
  },
  {
    code: "1238573915",
    name: "온더보더 광명에이케이점 (JRW)",
    channel: "체인",
    manager: "SY",
    storeType: "비매장",
    brand: "온더보더",
    status: "거래중",
  },
  {
    code: "8061202753",
    name: "버거넛",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "버거넛",
    status: "거래중",
  },
  {
    code: "8344800609",
    name: "힘난다버거 성복점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "힘난다(매장)",
    status: "거래중",
  },
  {
    code: "1201210765",
    name: "버거스하이",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "버거스하이",
    status: "거래중",
  },
  {
    code: "2093743295",
    name: "힘난다버거_광교점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "힘난다(매장)",
    status: "거래중",
  },
  {
    code: "3523301567",
    name: "치코스",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "치코스",
    status: "거래중",
  },
  {
    code: "3520403342",
    name: "백소정 주안아인병원점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "백소정",
    status: "거래중",
  },
  {
    code: "5571602276",
    name: "링키지 버거(LINKAGE BURGER)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "링키지 버거(LINKAGE BURGER)",
    status: "거래중",
  },
  {
    code: "6950503135",
    name: "니즈버거 여의도점",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "니즈버거 여의도점",
    status: "거래중",
  },
  {
    code: "7056100751",
    name: "바이트클럽(BITE CLUB)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "바이트클럽(BITE CLUB)",
    status: "거래중",
  },
  {
    code: "3266600698",
    name: "백소정 역삼GS점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "백소정",
    status: "거래중",
  },
  {
    code: "2368701424",
    name: "엘더버거",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "엘더버거",
    status: "거래중",
  },
  {
    code: "8020203407",
    name: "몽키필리",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "몽키필리",
    status: "거래중",
  },
  {
    code: "8276100217",
    name: "버거그랩(Burger grab)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "버거그랩(Burger grab)",
    status: "거래중",
  },
  {
    code: "5024638621",
    name: "백소정 도곡점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "백소정",
    status: "거래중",
  },
  {
    code: "1514200690",
    name: "포티데이즈",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "포티데이즈",
    status: "거래중",
  },
  {
    code: "5111066869",
    name: "스테이플버거 가산점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "스테이플버거 가산점",
    status: "거래중",
  },
  {
    code: "4350803456",
    name: "라카이(LAKAI)",
    channel: "권역배송",
    manager: "KT",
    storeType: "매장",
    brand: "라카이(LAKAI)",
    status: "거래중",
  },
  {
    code: "6475100586",
    name: "프레디버거 화곡역점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "프레디버거 화곡역점",
    status: "거래중",
  },
  {
    code: "1590502703",
    name: "백소정 방배역점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "백소정",
    status: "거래중",
  },
  {
    code: "2114601043",
    name: "경성버거",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "경성버거",
    status: "거래중",
  },
  {
    code: "8297700325",
    name: "19버거테이블",
    channel: "권역배송",
    manager: "KT",
    storeType: "매장",
    brand: "19버거테이블",
    status: "거래중",
  },
  {
    code: "7401902132",
    name: "디어버거 창원대점",
    channel: "권역배송",
    manager: "KT",
    storeType: "매장",
    brand: "디어버거",
    status: "거래중",
  },
  {
    code: "1420673959",
    name: "온더고(ONTHEGO)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "온더고(ONTHEGO)",
    status: "거래중",
  },
  {
    code: "7861402598",
    name: "노컷서울 금호",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "노컷서울 금호",
    status: "거래중",
  },
  {
    code: "5766500734",
    name: "버거 다이브",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "버거 다이브",
    status: "거래중",
  },
  {
    code: "6294200529",
    name: "꿈버거 상점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "꿈버거 상점",
    status: "거래중",
  },
  {
    code: "6780102735",
    name: "다원푸드",
    channel: "도매",
    manager: "NH",
    storeType: "매장",
    brand: "다원푸드",
    status: "거래중",
  },
  {
    code: "6407800490",
    name: "니즈버거 을지로점",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "니즈버거 을지로점",
    status: "거래중",
  },
  {
    code: "7942200640",
    name: "든해버거",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "든해버거",
    status: "거래중",
  },
  {
    code: "5994400969",
    name: "백소정 동양미래대점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "백소정",
    status: "거래중",
  },
  {
    code: "2445100878",
    name: "버문스버거 주안시민공원역점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "버문스버거 주안시민공원역점",
    status: "거래중",
  },
  {
    code: "1980802072",
    name: "보스턴수제버거 부평점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "보스턴수제버거 부평점",
    status: "거래중",
  },
  {
    code: "7061902199",
    name: "엉클캐빈",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "엉클캐빈",
    status: "거래중",
  },
  {
    code: "1461802390",
    name: "마시안 300",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "마시안 300",
    status: "거래중",
  },
  {
    code: "6490202807",
    name: "힘난다버거_양평역점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "힘난다(매장)",
    status: "거래중",
  },
  {
    code: "4403601486",
    name: "세포버거",
    channel: "권역배송",
    manager: "KT",
    storeType: "매장",
    brand: "세포버거",
    status: "거래중",
  },
  {
    code: "2301402381",
    name: "레이지오프(Lazy off)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "레이지오프(Lazy off)",
    status: "거래중",
  },
  {
    code: "2631802171",
    name: "케일럽버거",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "케일럽버거",
    status: "거래중",
  },
  {
    code: "5405200337",
    name: "티제이 버거앤 파스타",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "티제이 버거앤 파스타",
    status: "거래중",
  },
  {
    code: "2260672904",
    name: "버거웍스",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "버거웍스",
    status: "거래중",
  },
  {
    code: "2161072511",
    name: "힘난다버거_상암점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "힘난다(매장)",
    status: "거래중",
  },
  {
    code: "5293200875",
    name: "타운앤컨트리스햄버거스",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "타운앤컨트리스햄버거스",
    status: "거래중",
  },
  {
    code: "8195900585",
    name: "솜다리(솜다리 버거)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "솜다리(솜다리 버거)",
    status: "거래중",
  },
  {
    code: "2437700477",
    name: "조지아미고(georgiamigo)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "조지아미고(georgiamigo)",
    status: "거래중",
  },
  {
    code: "2171726339",
    name: "노컷서울 마곡",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "노컷서울 마곡",
    status: "거래중",
  },
  {
    code: "7380203519",
    name: "조선버거",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "조선버거",
    status: "거래중",
  },
  {
    code: "6566400805",
    name: "오버드라이브",
    channel: "권역배송",
    manager: "KT",
    storeType: "매장",
    brand: "오버드라이브",
    status: "거래중",
  },
  {
    code: "3598502552",
    name: "레츠잇치킨",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "레츠잇치킨",
    status: "거래중",
  },
  {
    code: "7382901726",
    name: "브라더s",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "브라더s",
    status: "거래중",
  },
  {
    code: "8378101910",
    name: "주식회사 호맥",
    channel: "제조",
    manager: "KT",
    storeType: "비매장",
    brand: "주식회사 호맥",
    status: "거래중",
  },
  {
    code: "8106600761",
    name: "버거리 인천 루원시티점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "버거리 인천 루원시티점",
    status: "거래중",
  },
  {
    code: "4012471719",
    name: "미분당 영종도점",
    channel: "체인",
    manager: "NH",
    storeType: "매장",
    brand: "미분당 영종도점",
    status: "거래중",
  },
  {
    code: "1164401327",
    name: "노스트레스버거 한남점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "노스트레스버거 한남점",
    status: "거래중",
  },
  {
    code: "5492401932",
    name: "노스트레스버거 신흥",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "노스트레스버거 신흥",
    status: "거래중",
  },
  {
    code: "2890403971",
    name: "번앤샷 BURN&SHOT",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "번앤샷 BURN&SHOT",
    status: "거래중",
  },
  {
    code: "5058528664",
    name: "프라이도화",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "프라이도화",
    status: "거래중",
  },
  {
    code: "6221802287",
    name: "슬로우키친",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "슬로우키친",
    status: "거래중",
  },
  {
    code: "2780702750",
    name: "비스티버거 반포직영",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "비스티버거 반포직영",
    status: "거래중",
  },
  {
    code: "6975300402",
    name: "브로버거(방이점)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "브로버거(방이점)",
    status: "거래중",
  },
  {
    code: "8030101083",
    name: "더브라이언스",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "더브라이언스",
    status: "거래중",
  },
  {
    code: "4291301990",
    name: "진주 수제버거 LAKAI",
    channel: "권역배송",
    manager: "KT",
    storeType: "매장",
    brand: "진주 수제버거 LAKAI",
    status: "거래중",
  },
  {
    code: "3021192722",
    name: "성수망치버거",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "성수망치버거",
    status: "거래중",
  },
  {
    code: "3393801123",
    name: "곰버거",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "곰버거",
    status: "거래중",
  },
  {
    code: "1981901860",
    name: "미분당 송내점",
    channel: "체인",
    manager: "NH",
    storeType: "매장",
    brand: "미분당 송내점",
    status: "거래중",
  },
  {
    code: "5054800899",
    name: "푸라닭 옥수점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "푸라닭 옥수점",
    status: "거래중",
  },
  {
    code: "6154700689",
    name: "푸라닭 정릉1호점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "푸라닭 정릉1호점",
    status: "거래중",
  },
  {
    code: "4548502255",
    name: "콩킨누스",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "콩킨누스",
    status: "거래중",
  },
  {
    code: "6481302528",
    name: "우프스낵바",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "우프스낵바",
    status: "거래중",
  },
  {
    code: "3490203037",
    name: "오엑스 그릴 OXG 도곡점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "오엑스 그릴 OXG 도곡점",
    status: "거래중",
  },
  {
    code: "8973001225",
    name: "그랜드버거",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "그랜드버거",
    status: "거래중",
  },
  {
    code: "1772901519",
    name: "글래디버거(이천중리지구점)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "글래디버거(이천중리지구점)",
    status: "거래중",
  },
  {
    code: "5440403091",
    name: "버거스타디움(BURGERSTADIUM)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "버거스타디움(BURGERSTADIUM)",
    status: "거래중",
  },
  {
    code: "2742302025",
    name: "썬이스트 버거",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "썬이스트 버거",
    status: "거래중",
  },
  {
    code: "4691202331",
    name: "베어스타코 용인성복점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "베어스타코 용인성복점",
    status: "거래중",
  },
  {
    code: "3223601835",
    name: "패티스버거 사당점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "패티스버거 사당점",
    status: "거래중",
  },
  {
    code: "3015900861",
    name: "버거플리즈",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "버거플리즈",
    status: "거래중",
  },
  {
    code: "1435700112",
    name: "싼타클로스(산타버거)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "싼타클로스(산타버거)",
    status: "거래중",
  },
  {
    code: "361-85-02609",
    name: "제스티살룬 목동현대점",
    channel: "체인",
    manager: "SW",
    storeType: "매장",
    brand: "제스티살룬",
    status: "거래중",
  },
  {
    code: "5178802237",
    name: "(주)제이에스에프엔비",
    channel: "도매",
    manager: "NH",
    storeType: "비매장",
    brand: "미인계",
    status: "거래중",
  },
  {
    code: "7792801604",
    name: "푸라닭 묵동점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "푸라닭 묵동점",
    status: "거래중",
  },
  {
    code: "3325800908",
    name: "비기버디",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "비기버디",
    status: "거래중",
  },
  {
    code: "4505800507",
    name: "풀리너마이트_홍대본점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "풀리너마이트_홍대본점",
    status: "거래중",
  },
  {
    code: "2526700375",
    name: "니즈(NEEDS)",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "니즈버거 신촌점",
    status: "거래중",
  },
  {
    code: "4012354557",
    name: "레서",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "레서",
    status: "거래중",
  },
  {
    code: "7265000941",
    name: "백소정 약수역점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "백소정",
    status: "거래중",
  },
  {
    code: "5080476676",
    name: "힘난다버거_안동옥동점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "힘난다(매장)",
    status: "거래중",
  },
  {
    code: "1213634427",
    name: "미국버거_봉덕점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "미국버거_봉덕점",
    status: "거래중",
  },
  {
    code: "4492002123",
    name: "프랭크버거 의정부금오신곡점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "프랭크버거 의정부금오신곡점",
    status: "거래중",
  },
  {
    code: "1487600307",
    name: "메이플 버거",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "메이플 버거",
    status: "거래중",
  },
  {
    code: "3698702738",
    name: "주식회사 유안아이앤씨",
    channel: "도매",
    manager: "SY",
    storeType: "비매장",
    brand: "주식회사 유안아이앤씨",
    status: "거래중",
  },
  {
    code: "6848500000",
    name: "후참잘 (하이닥)",
    channel: "체인",
    manager: "SY",
    storeType: "비매장",
    brand: "후라이드참잘하는집",
    status: "거래중",
  },
  {
    code: "1390282890",
    name: "은진유통",
    channel: "도매",
    manager: "SY",
    storeType: "비매장",
    brand: "은진유통",
    status: "거래중",
  },
  {
    code: "6845800003",
    name: "후참잘 (SFN)",
    channel: "체인",
    manager: "SY",
    storeType: "비매장",
    brand: "후라이드참잘하는집",
    status: "거래중",
  },
  {
    code: "6848500001",
    name: "후참잘 (도한센터)",
    channel: "체인",
    manager: "SY",
    storeType: "비매장",
    brand: "후라이드참잘하는집",
    status: "거래중",
  },
  {
    code: "4442501667",
    name: "보스턴수제버거",
    channel: "체인",
    manager: "NH",
    storeType: "매장",
    brand: "보스턴수제버거",
    status: "거래중",
  },
  {
    code: "1138106497",
    name: "(주) 조흥",
    channel: "도매",
    manager: "SY",
    storeType: "비매장",
    brand: "(주) 조흥",
    status: "거래중",
  },
  {
    code: "6921700277",
    name: "예스한샘유통",
    channel: "체인물류",
    manager: "SY",
    storeType: "비매장",
    brand: "예스한샘유통",
    status: "거래중",
  },
  {
    code: "7438100222",
    name: "(주)힘난다",
    channel: "체인",
    manager: "SY",
    storeType: "비매장",
    brand: "힘난다(본사)",
    status: "거래중",
  },
  {
    code: "4678601074-1",
    name: "세이웰_375브런치(아워홈)",
    channel: "체인물류",
    manager: "SY",
    storeType: "비매장",
    brand: "세이웰",
    status: "거래중",
  },
  {
    code: "2103951708",
    name: "잭스패티",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "잭스패티",
    status: "거래중",
  },
  {
    code: "5210303451",
    name: "치벅",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "치벅",
    status: "거래중",
  },
  {
    code: "6595400751",
    name: "힘난다버거_인천마전점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "힘난다(매장)",
    status: "거래중",
  },
  {
    code: "2031312217",
    name: "미분당 배곧신도시점",
    channel: "체인",
    manager: "NH",
    storeType: "매장",
    brand: "미분당 배곧신도시점",
    status: "거래중",
  },
  {
    code: "522-11-02673",
    name: "메이상사",
    channel: "도매",
    manager: "SY",
    storeType: "비매장",
    brand: "이마트 (수출)",
    status: "거래중",
  },
  {
    code: "5698603586",
    name: "주식회사 아란치니브라더스",
    channel: "제조",
    manager: "SY",
    storeType: "비매장",
    brand: "주식회사 아란치니브라더스",
    status: "거래중",
  },
  {
    code: "7398800843",
    name: "(주)한승인터내쇼날",
    channel: "도매",
    manager: "SY",
    storeType: "비매장",
    brand: "(주)한승인터내쇼날",
    status: "거래중",
  },
  {
    code: "6038111645",
    name: "CJFW(주)안양센터",
    channel: "기업",
    manager: "SY",
    storeType: "비매장",
    brand: "CJFW",
    status: "거래중",
  },
  {
    code: "121-81-40026",
    name: "주식회사 한국도매물류",
    channel: "도매",
    manager: "SY",
    storeType: "비매장",
    brand: "주식회사 한국도매물류",
    status: "거래중",
  },
  {
    code: "8031601675",
    name: "한소쿠리",
    channel: "도매",
    manager: "SY",
    storeType: "비매장",
    brand: "한소쿠리",
    status: "거래중",
  },
  {
    code: "3878102500",
    name: "(주)인맥-올잇마켓",
    channel: "체인",
    manager: "SY",
    storeType: "비매장",
    brand: "(주)인맥에프엔씨",
    status: "거래중",
  },
  {
    code: "7520203868",
    name: "이븐버거(EVEN BURGER)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "이븐버거",
    status: "거래중",
  },
  {
    code: "5621202038",
    name: "초가삼간 보스턴수제버거",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "초가삼간 보스턴수제버거",
    status: "거래중",
  },
  {
    code: "6138700700",
    name: "농업회사법인 우리쌀푸드(주)",
    channel: "제조",
    manager: "NH",
    storeType: "비매장",
    brand: "농업회사법인 우리쌀푸드(주)",
    status: "거래중",
  },
  {
    code: "4611602655",
    name: "치벅(CHEEBUCK) 강남구청점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "치벅(CHEEBUCK) 강남구청점",
    status: "거래중",
  },
  {
    code: "765-07-01968",
    name: "플레이벅",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "플레이벅",
    status: "거래중",
  },
  {
    code: "4108700697",
    name: "주식회사 월광식자재",
    channel: "식자재마트",
    manager: "KT",
    storeType: "비매장",
    brand: "주식회사 월광식자재",
    status: "거래중",
  },
  {
    code: "3082901715",
    name: "버거플리",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "버거플리",
    status: "거래중",
  },
  {
    code: "891881802091",
    name: "주식회사 에이비랩코리아",
    channel: "본사",
    manager: "Bomi",
    storeType: "비매장",
    brand: "주식회사 에이비랩코리아",
    status: "거래중",
  },
  {
    code: "124-86-61480",
    name: "가나식품주식회사",
    channel: "매입",
    manager: "Bomi",
    storeType: "비매장",
    brand: "가나식품주식회사",
    status: "거래중",
  },
  {
    code: "에이비랩",
    name: "직원구매",
    channel: "본사",
    manager: "Bomi",
    storeType: "비매장",
    brand: "직원구매",
    status: "거래중",
  },
  {
    code: "샘플",
    name: "에이비랩_샘플",
    channel: "본사",
    manager: "Bomi",
    storeType: "비매장",
    brand: "에이비랩_샘플",
    status: "거래중",
  },
  {
    code: "1078112091",
    name: "제니코식품 (주)",
    channel: "매입",
    manager: "Bomi",
    storeType: "비매장",
    brand: "제니코식품 (주)",
    status: "거래중",
  },
  {
    code: "7862702017",
    name: "비케이펍앤버거(BKPUP& Burger)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "비케이펍앤버거(BKPUP& Burger)",
    status: "거래종료",
  },
  {
    code: "3396400711",
    name: "세븐패티버거 가산점",
    channel: "체인",
    manager: "SW",
    storeType: "매장",
    brand: "세븐패티버거",
    status: "거래종료",
  },
  {
    code: "5500103678",
    name: "세븐패티버거 죽전점",
    channel: "체인",
    manager: "SW",
    storeType: "매장",
    brand: "세븐패티버거",
    status: "거래종료",
  },
  {
    code: "748-86-02700",
    name: "제스티살룬 파주운정점",
    channel: "체인",
    manager: "SW",
    storeType: "매장",
    brand: "제스티살룬",
    status: "거래종료",
  },
  {
    code: "3198701393",
    name: "주식회사 더백에프앤비",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "주식회사 더백에프앤비",
    status: "거래종료",
  },
  {
    code: "3198701302",
    name: "더백푸드트럭본점(해방촌점)",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "더백푸드트럭본점(해방촌점)",
    status: "거래종료",
  },
  {
    code: "6822401113",
    name: "제레미버거_도산점",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "제레미버거_도산점",
    status: "거래종료",
  },
  {
    code: "2688802994",
    name: "주식회사 오오티오",
    channel: "체인",
    manager: "SW",
    storeType: "매장",
    brand: "세븐패티버거",
    status: "거래종료",
  },
  {
    code: "2580403069",
    name: "버거타임",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "버거타임",
    status: "거래종료",
  },
  {
    code: "6840501122",
    name: "노컷서울 금호",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "노컷서울 금호",
    status: "거래종료",
  },
  {
    code: "8585900862",
    name: "도그즈인번즈(Dogs Buns)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "도그즈인번즈(Dogs Buns)",
    status: "거래종료",
  },
  {
    code: "2073161771",
    name: "패티패티",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "패티패티",
    status: "거래종료",
  },
  {
    code: "4368601065",
    name: "웰빙푸드시스템 주식회사",
    channel: "체인",
    manager: "SY",
    storeType: "비매장",
    brand: "666버거",
    status: "거래종료",
  },
  {
    code: "5641902066",
    name: "에이스버거 동대문 본점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "에이스버거 동대문 본점",
    status: "거래종료",
  },
  {
    code: "6305200912",
    name: "브릿지헤드코리아",
    channel: "도매",
    manager: "SY",
    storeType: "비매장",
    brand: "브릿지헤드코리아",
    status: "거래종료",
  },
  {
    code: "1238574777",
    name: "프레시원주식회사 남서울사업부",
    channel: "기업",
    manager: "SY",
    storeType: "비매장",
    brand: "CJFW",
    status: "거래종료",
  },
  {
    code: "3198722222",
    name: "주식회사 더백에프앤비(더현대서울점)",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "주식회사 더백에프앤비(더현대서울점)",
    status: "거래종료",
  },
  {
    code: "4308503068",
    name: "테이스티버거 대전둔산점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "테이스티버거 대전둔산점",
    status: "거래종료",
  },
  {
    code: "1298688941",
    name: "(주)에스피씨 지에프에스_피자와썹",
    channel: "체인",
    manager: "KT",
    storeType: "비매장",
    brand: "SPC(피자와썹)",
    status: "거래종료",
  },
  {
    code: "8700902429",
    name: "바운스무드 오창점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "바운스무드 오창점",
    status: "거래종료",
  },
  {
    code: "8631202611",
    name: "바운스무드 복대점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "바운스무드 복대점",
    status: "거래종료",
  },
  {
    code: "2178124157",
    name: "다인에프씨 주식회사",
    channel: "도매",
    manager: "SY",
    storeType: "비매장",
    brand: "다인에프씨 주식회사",
    status: "거래종료",
  },
  {
    code: "5044114799",
    name: "유피버거",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "유피버거",
    status: "거래종료",
  },
  {
    code: "2378802985-1",
    name: "주식회사 하라에프에스 (버거옥)",
    channel: "도매",
    manager: "SY",
    storeType: "비매장",
    brand: "주식회사 하라에프에스 (버거옥)",
    status: "거래종료",
  },
  {
    code: "1268639984",
    name: "농업회사법인 태성그린푸드 주식회사",
    channel: "체인물류",
    manager: "SY",
    storeType: "비매장",
    brand: "닭장수후라이드",
    status: "거래종료",
  },
  {
    code: "1508700999",
    name: "(주)비스티보이즈(비스티버거)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "(주)비스티보이즈(비스티버거)",
    status: "거래종료",
  },
  {
    code: "1750702484",
    name: "백소정 검단사거리역점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "백소정",
    status: "거래종료",
  },
  {
    code: "6524301127",
    name: "홀리몰리버거_창원",
    channel: "권역배송",
    manager: "KT",
    storeType: "매장",
    brand: "홀리몰리버거_창원",
    status: "거래종료",
  },
  {
    code: "2178124157-1",
    name: "다인에프씨 주식회사 (다인식품)",
    channel: "도매",
    manager: "SY",
    storeType: "비매장",
    brand: "다인에프씨 주식회사 (다인식품)",
    status: "거래종료",
  },
  {
    code: "7158502918",
    name: "주식회사 지원글로벌 백소정 계양구청점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "백소정",
    status: "거래종료",
  },
  {
    code: "2011014519",
    name: "더백푸드트럭본점(해방촌점)",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "더백푸드트럭본점(해방촌점)",
    status: "거래종료",
  },
  {
    code: "1332002548",
    name: "포테이토헤드 버거샵",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "포테이토헤드 버거샵",
    status: "거래종료",
  },
  {
    code: "2211562422",
    name: "테이스티버거2018",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "테이스티버거2018",
    status: "거래종료",
  },
  {
    code: "7528802086",
    name: "에이에프컴퍼니",
    channel: "도매",
    manager: "NH",
    storeType: "비매장",
    brand: "에이에프컴퍼니",
    status: "거래종료",
  },
  {
    code: "2964601030",
    name: "프레클버거(Freckle Burger)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "프레클버거(Freckle Burger)",
    status: "거래종료",
  },
  {
    code: "2655501076",
    name: "파파로카",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "파파로카",
    status: "거래종료",
  },
  {
    code: "6473101215",
    name: "보스턴 영월",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "보스턴 영월",
    status: "거래종료",
  },
  {
    code: "7263401009",
    name: "니즈버거 청라점",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "니즈버거 청라점",
    status: "거래종료",
  },
  {
    code: "3437500356",
    name: "푸라닭 신내점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "푸라닭 신내점",
    status: "거래종료",
  },
  {
    code: "8484700668",
    name: "알지비버거(rgbburger)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "알지비버거(rgbburger)",
    status: "거래종료",
  },
  {
    code: "2703201593",
    name: "홀리몰리버거(HOLYMOLYBURGER)_진해점",
    channel: "권역배송",
    manager: "KT",
    storeType: "매장",
    brand: "홀리몰리버거(HOLYMOLYBURGER)_진해점",
    status: "거래종료",
  },
  {
    code: "3266200685",
    name: "19버거테이블(부산시청점)",
    channel: "권역배송",
    manager: "KT",
    storeType: "매장",
    brand: "19버거테이블_부산시청점",
    status: "거래종료",
  },
  {
    code: "2791900914",
    name: "분더버거",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "분더버거",
    status: "거래종료",
  },
  {
    code: "1853501369",
    name: "힘난다버거(코엑스점)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "힘난다버거(코엑스점)",
    status: "거래종료",
  },
  {
    code: "8681602633",
    name: "브로버거(광명점)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "브로버거(광명점)",
    status: "거래종료",
  },
  {
    code: "6644700696",
    name: "버거106(BURGER106)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "버거106",
    status: "거래종료",
  },
  {
    code: "7528601875",
    name: "(주)지인프라임",
    channel: "체인",
    manager: "SY",
    storeType: "비매장",
    brand: "(주)지인프라임",
    status: "거래종료",
  },
  {
    code: "4268702967",
    name: "백소정 부천역점(㈜다원글로벌)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "백소정",
    status: "거래종료",
  },
  {
    code: "1634701208",
    name: "버거치즈스마일",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "버거치즈스마일",
    status: "거래종료",
  },
  {
    code: "6121893990",
    name: "백소정 철산점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "백소정",
    status: "거래종료",
  },
  {
    code: "8938502969",
    name: "힘난다버거 부천상동점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "힘난다(매장)",
    status: "거래종료",
  },
  {
    code: "1028508789",
    name: "백소정 주안역점(㈜다원글로벌)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "백소정",
    status: "거래종료",
  },
  {
    code: "4581801538",
    name: "미분당 계양구청점",
    channel: "체인",
    manager: "NH",
    storeType: "매장",
    brand: "미분당 계양구청점",
    status: "거래종료",
  },
  {
    code: "7262501026",
    name: "미분당",
    channel: "체인",
    manager: "NH",
    storeType: "매장",
    brand: "미분당 청라점",
    status: "거래종료",
  },
  {
    code: "4728502796",
    name: "(주)지원글로벌 백소정역곡역점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "백소정",
    status: "거래종료",
  },
  {
    code: "8068603205",
    name: "(주)양지바른",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "(주)양지바른",
    status: "거래종료",
  },
  {
    code: "8795500764",
    name: "디어버거 창원가로수길점 (본점)",
    channel: "권역배송",
    manager: "KT",
    storeType: "매장",
    brand: "디어버거",
    status: "거래종료",
  },
  {
    code: "3793001014",
    name: "에프터드링크버거",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "에프터드링크버거",
    status: "거래종료",
  },
  {
    code: "2118149636",
    name: "(주) 제때_서가앤쿡",
    channel: "도매",
    manager: "KT",
    storeType: "비매장",
    brand: "(주) 제때_서가앤쿡",
    status: "거래종료",
  },
  {
    code: "5061711435",
    name: "오대양푸드",
    channel: "도매",
    manager: "KT",
    storeType: "비매장",
    brand: "오대양푸드",
    status: "거래종료",
  },
  {
    code: "6623401501",
    name: "디어버거 갤러리아점",
    channel: "권역배송",
    manager: "KT",
    storeType: "매장",
    brand: "디어버거",
    status: "거래종료",
  },
  {
    code: "1160819585",
    name: "끄트머리피자 (디어버거 사천2호점)",
    channel: "권역배송",
    manager: "KT",
    storeType: "매장",
    brand: "디어버거",
    status: "거래종료",
  },
  {
    code: "1888100196",
    name: "주식회사 중앙에프엔에스",
    channel: "식자재마트",
    manager: "KT",
    storeType: "비매장",
    brand: "주식회사 중앙에프엔에스",
    status: "거래종료",
  },
  {
    code: "1088176929",
    name: "주식회사 제이알더블유",
    channel: "체인",
    manager: "SY",
    storeType: "비매장",
    brand: "온더보더",
    status: "거래종료",
  },
  {
    code: "2158679307",
    name: "유한회사 현아농산",
    channel: "도매",
    manager: "SY",
    storeType: "비매장",
    brand: "유한회사 현아농산",
    status: "거래종료",
  },
  {
    code: "397-86-03531",
    name: "주식회사 성신",
    channel: "제조",
    manager: "SY",
    storeType: "비매장",
    brand: "주식회사 성신",
    status: "거래종료",
  },
  {
    code: "2757700606",
    name: "태광유통",
    channel: "도매",
    manager: "SY",
    storeType: "비매장",
    brand: "태광유통",
    status: "거래종료",
  },
  {
    code: "1408170064",
    name: "주식회사 건우유통",
    channel: "도매",
    manager: "SY",
    storeType: "비매장",
    brand: "주식회사 건우유통",
    status: "거래종료",
  },
  {
    code: "1068514847",
    name: "해태제과식품 (주) 서울지점",
    channel: "도매",
    manager: "SY",
    storeType: "비매장",
    brand: "해태제과식품 (주) 서울지점",
    status: "거래종료",
  },
  {
    code: "2030431454",
    name: "코코푸드",
    channel: "체인물류",
    manager: "SY",
    storeType: "비매장",
    brand: "코코푸드",
    status: "거래종료",
  },
  {
    code: "1598603109",
    name: "주식회사 사우스코어",
    channel: "제조",
    manager: "SY",
    storeType: "비매장",
    brand: "주식회사 사우스코어",
    status: "거래종료",
  },
  {
    code: "3198733333",
    name: "주식회사 더백에프앤비(작업장)",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "주식회사 더백에프앤비(작업장)",
    status: "거래종료",
  },
  {
    code: "3512602100",
    name: "코코벅",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "코코벅",
    status: "거래종료",
  },
  {
    code: "4263900390",
    name: "에이플렛(A_plat)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "에이플랫(A_PLAT) 버거",
    status: "거래종료",
  },
  {
    code: "5383201387",
    name: "캐비넷베이글",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "캐비넷 베이글",
    status: "거래종료",
  },
  {
    code: "7578802325",
    name: "주식회사 두잇(Doeat)",
    channel: "제조",
    manager: "NH",
    storeType: "비매장",
    brand: "주식회사 두잇(Doeat)",
    status: "거래종료",
  },
  {
    code: "8348502602",
    name: "주식회사 누땡",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "주식회사 누땡",
    status: "거래종료",
  },
  {
    code: "5053510081",
    name: "백소정 판교점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "백소정",
    status: "거래종료",
  },
  {
    code: "2824101088",
    name: "만나푸드",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "만나푸드",
    status: "거래종료",
  },
  {
    code: "2136700797",
    name: "바운스무드 동남지구점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "바운스무드 동남지구점",
    status: "거래종료",
  },
  {
    code: "1683701669",
    name: "바운스무드 청주 본점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "바운스무드 청주 본점",
    status: "거래종료",
  },
  {
    code: "7886400795",
    name: "아지트카페앤펍(CAFE&PUB)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "아지트카페앤펍(CAFE&PUB)",
    status: "거래종료",
  },
  {
    code: "2653300814",
    name: "스매쉬드 매스",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "스매쉬드 매스",
    status: "거래종료",
  },
  {
    code: "1318802877",
    name: "주식회사 포코너스",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "주식회사 포코너스",
    status: "거래종료",
  },
  {
    code: "8150402427",
    name: "원더샐러드(힘난다버거 코엑스점)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "원더샐러드(힘난다버거 코엑스점)",
    status: "거래종료",
  },
  {
    code: "4652201891",
    name: "아미르 케밥(AMLR KEBAB)",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "아미르 케밥(AMLR KEBAB)",
    status: "거래종료",
  },
  {
    code: "0014",
    name: "브루클린버거더조인트 청계천점",
    channel: "체인",
    manager: "NH",
    storeType: "비매장",
    brand: "브루클린",
    status: "거래종료",
  },
  {
    code: "0018",
    name: "브루클린더버거조인트 판교점",
    channel: "체인",
    manager: "NH",
    storeType: "비매장",
    brand: "브루클린",
    status: "거래종료",
  },
  {
    code: "0022",
    name: "브루클린더버거조인트 올림픽공원점",
    channel: "체인",
    manager: "NH",
    storeType: "비매장",
    brand: "브루클린",
    status: "거래종료",
  },
  {
    code: "0010",
    name: "브루클린더버거조인트 여의도점",
    channel: "체인",
    manager: "NH",
    storeType: "비매장",
    brand: "브루클린",
    status: "거래종료",
  },
  {
    code: "0013",
    name: "브루클린더버거조인트 분당정자점",
    channel: "체인",
    manager: "NH",
    storeType: "비매장",
    brand: "브루클린",
    status: "거래종료",
  },
  {
    code: "0016",
    name: "브루클린더버거조인트 목동점",
    channel: "체인",
    manager: "NH",
    storeType: "비매장",
    brand: "브루클린",
    status: "거래종료",
  },
  {
    code: "0011",
    name: "브루클린더버거조인트 롯데월드몰점",
    channel: "체인",
    manager: "NH",
    storeType: "비매장",
    brand: "브루클린",
    status: "거래종료",
  },
  {
    code: "0009",
    name: "브루클린더버거조인트 동부이촌점",
    channel: "체인",
    manager: "NH",
    storeType: "비매장",
    brand: "브루클린",
    status: "거래종료",
  },
  {
    code: "0023",
    name: "브루클린 더 버거 조인트 역삼점",
    channel: "체인",
    manager: "NH",
    storeType: "비매장",
    brand: "브루클린",
    status: "거래종료",
  },
  {
    code: "5868503003",
    name: "(주)임어전씨 수제버거 전곡역지점",
    channel: "체인",
    manager: "NH",
    storeType: "매장",
    brand: "(주)임어전씨 수제버거 전곡역지점",
    status: "거래종료",
  },
  {
    code: "0003",
    name: "(주)에스씨비에이치 서래마을점",
    channel: "체인",
    manager: "NH",
    storeType: "비매장",
    brand: "브루클린",
    status: "거래종료",
  },
  {
    code: "0004",
    name: "(주)에스씨비에이치 삼성점",
    channel: "체인",
    manager: "NH",
    storeType: "비매장",
    brand: "브루클린",
    status: "거래종료",
  },
  {
    code: "0008",
    name: "(주)에스씨비에이치 광화문 디타워점",
    channel: "체인",
    manager: "NH",
    storeType: "비매장",
    brand: "브루클린",
    status: "거래종료",
  },
  {
    code: "0002",
    name: "(주)에스씨비에이치 가로수길점",
    channel: "체인",
    manager: "NH",
    storeType: "비매장",
    brand: "브루클린",
    status: "거래종료",
  },
  {
    code: "3134500718",
    name: "페어쉬림프세라믹",
    channel: "도매",
    manager: "NH",
    storeType: "비매장",
    brand: "페어쉬림프세라믹",
    status: "거래종료",
  },
  {
    code: "5900503179",
    name: "힘난다버거 송도점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "힘난다(매장)",
    status: "거래종료",
  },
  {
    code: "3061451272",
    name: "피자탑 포항남구점(블럭키친 포항남구점)",
    channel: "권역배송",
    manager: "KT",
    storeType: "매장",
    brand: "블럭키친",
    status: "거래종료",
  },
  {
    code: "1691901106",
    name: "미스터키친 (블럭키친 경남고성점)",
    channel: "권역배송",
    manager: "KT",
    storeType: "매장",
    brand: "블럭키친",
    status: "거래종료",
  },
  {
    code: "6108702830",
    name: "주식회사 이음에프엔아이",
    channel: "도매",
    manager: "KT",
    storeType: "비매장",
    brand: "주식회사 이음에프엔아이",
    status: "거래종료",
  },
  {
    code: "1063301362",
    name: "제일유통",
    channel: "도매",
    manager: "KT",
    storeType: "비매장",
    brand: "제일유통",
    status: "거래종료",
  },
  {
    code: "3948702433",
    name: "이음푸드",
    channel: "도매",
    manager: "KT",
    storeType: "비매장",
    brand: "이음푸드",
    status: "거래종료",
  },
  {
    code: "3128125280",
    name: "(주)동원홈푸드",
    channel: "도매",
    manager: "KT",
    storeType: "비매장",
    brand: "(주)동원홈푸드",
    status: "거래종료",
  },
  {
    code: "4555101045",
    name: "선산맛닭",
    channel: "권역배송",
    manager: "KT",
    storeType: "매장",
    brand: "선산맛닭",
    status: "거래종료",
  },
  {
    code: "3918502857",
    name: "프레시원주식회사 부산사업부",
    channel: "기업",
    manager: "KT",
    storeType: "비매장",
    brand: "CJFW",
    status: "거래종료",
  },
  {
    code: "7291502248",
    name: "조선버거",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "조선버거",
    status: "거래종료",
  },
  {
    code: "4821501961",
    name: "행루즈버거",
    channel: "권역배송",
    manager: "SW",
    storeType: "매장",
    brand: "행루즈버거",
    status: "거래종료",
  },
  {
    code: "7320402743",
    name: "테이스티버거 대전둔산점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "테이스티버거 대전둔산점",
    status: "거래종료",
  },
  {
    code: "5935700879",
    name: "버거플리",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "버거플리",
    status: "거래종료",
  },
  {
    code: "5221202516",
    name: "백소정 서창점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "백소정",
    status: "거래종료",
  },
  {
    code: "1051996953",
    name: "버거넛",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "버거넛",
    status: "거래종료",
  },
  {
    code: "4663501287",
    name: "백소정 도곡점",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "백소정",
    status: "거래종료",
  },
  {
    code: "6371902253",
    name: "오버드라이브",
    channel: "권역배송",
    manager: "KT",
    storeType: "매장",
    brand: "오버드라이브",
    status: "거래종료",
  },
  {
    code: "1772601916",
    name: "버거바",
    channel: "권역배송",
    manager: "NH",
    storeType: "매장",
    brand: "버거바",
    status: "거래종료",
  },
];

const initialTargets: TargetRecord[] = [
  { storeType: "매장", month: thisMonth(), amount: 20000000 },
  { storeType: "비매장", month: thisMonth(), amount: 15000000 },
];

const initialEsts: EstRecord[] = [
  {
    storeCode: "A001",
    storeName: "강남점",
    month: thisMonth(),
    amount: 18000000,
  },
  {
    storeCode: "A002",
    storeName: "홍대점",
    month: thisMonth(),
    amount: 14000000,
  },
];

const initialItemCosts: ItemCostRecord[] = [];

const initialSales: SalesRecord[] = [
  makeSale(
    "current",
    thisMonth(),
    `${thisMonth()}-17`,
    "A001",
    "강남점",
    "P001",
    "샘플상품",
    1,
    12500000,
    8000000,
    4500000,
    initialStores,
  ),
  makeSale(
    "current",
    thisMonth(),
    `${thisMonth()}-17`,
    "A002",
    "홍대점",
    "P001",
    "샘플상품",
    1,
    5300000,
    3500000,
    1800000,
    initialStores,
  ),
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
  const diff = Math.floor(
    (to.getTime() - from.getTime()) / (1000 * 60 * 60 * 24),
  );
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

function displayBrand(v: unknown) {
  const brand = norm(v);
  return !brand || brand === "미지정" ? "당월 신규 거래처" : brand;
}

function normalizeStatus(v: unknown): Store["status"] {
  const t = norm(v);
  return t === "종료" ||
    t === "거래종료" ||
    t === "비활성" ||
    t === "비활성화" ||
    t.toLowerCase() === "inactive" ||
    t.toLowerCase() === "closed"
    ? "거래종료"
    : "거래중";
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
    if (d)
      return `${d.y}-${String(d.m).padStart(2, "0")}-${String(d.d).padStart(2, "0")}`;
  }
  const t = norm(v).replaceAll(".", "-").replaceAll("/", "-");
  if (/^\d{8}$/.test(t))
    return `${t.slice(0, 4)}-${t.slice(4, 6)}-${t.slice(6, 8)}`;
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
  if (day === 0 || day === 6) return 0.5;
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

  for (
    let d = start, guard = 0;
    d <= end && guard < 40;
    d = addDays(d, 1), guard += 1
  ) {
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

function resolveStoreInfo(
  storeCode: string,
  fallbackName: string,
  fallback: Partial<Store>,
  stores: Store[],
) {
  const original = storeMap(stores).get(storeCode);

  return {
    code: storeCode || fallbackName || "미지정",
    name: original?.name || fallbackName || storeCode || "미지정",
    channel: original?.channel || fallback.channel || "미지정",
    manager: original?.manager || fallback.manager || "미지정",
    storeType: original?.storeType || fallback.storeType || "비매장",
    brand: displayBrand(original?.brand || fallback.brand),
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
  uploadedProfitRate?: number,
): SalesRecord {
  const s = storeMap(stores).get(storeCode);
  const profitRate = Number.isFinite(uploadedProfitRate)
    ? Number(uploadedProfitRate)
    : salesAmount
      ? (profitAmount / salesAmount) * 100
      : 0;
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
    brand: displayBrand(s?.brand),
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


const STATIC_SYNC_INTERVAL_MS = 30 * 60 * 1000;
const SALES_SYNC_INTERVAL_MS = 5 * 60 * 1000;
const FOCUS_REFRESH_MIN_GAP_MS = 5 * 60 * 1000;
const MAX_SUPABASE_BACKOFF_MS = 30 * 60 * 1000;

let supabaseBackoffUntil = 0;
let supabaseFailureCount = 0;

function canCallSupabase() {
  return Date.now() >= supabaseBackoffUntil;
}

function registerSupabaseSuccess() {
  supabaseFailureCount = 0;
  supabaseBackoffUntil = 0;
}

function registerSupabaseFailure(status: number) {
  if (status !== 429 && status !== 500 && status !== 502 && status !== 503 && status !== 504) return;
  supabaseFailureCount += 1;
  const delay = Math.min(30_000 * 2 ** Math.min(supabaseFailureCount - 1, 6), MAX_SUPABASE_BACKOFF_MS);
  supabaseBackoffUntil = Date.now() + delay;
}

function sharedApiConfig() {
  return {
    endpoint: "/api/d1/settings",
    headers: { "Content-Type": "application/json" },
  };
}

async function loadSharedStateRow<T>(
  key: string,
): Promise<AppStateRow<T> | null> {
  const config = sharedApiConfig();
  const response = await fetch(
    `${config.endpoint}/${encodeURIComponent(key)}`,
    { method: "GET", headers: config.headers, cache: "no-store" },
  );

  if (response.status === 404) return null;
  if (!response.ok) throw new Error(`D1 load failed: ${response.status}`);
  return (await response.json()) as AppStateRow<T>;
}

async function loadSharedState<T>(key: string): Promise<T | null> {
  const row = await loadSharedStateRow<T>(key);
  return row?.data ?? null;
}

function localMetaKey(key: string) {
  return `${key}__local_meta`;
}


function safeSetLocalStorage(key: string, value: string) {
  if (typeof window === "undefined") return false;
  try {
    window.localStorage.setItem(key, value);
    return true;
  } catch (error) {
    console.warn(
      "브라우저 저장소 용량이 부족해 localStorage 저장을 건너뜁니다. Cloudflare D1 저장은 계속 시도합니다.",
      error,
    );
    return false;
  }
}

function safeGetLocalStorage(key: string) {
  if (typeof window === "undefined") return null;
  try {
    return window.localStorage.getItem(key);
  } catch (error) {
    console.warn("브라우저 저장소 읽기 실패", error);
    return null;
  }
}

function getLocalMeta(key: string) {
  if (typeof window === "undefined") return { editedAt: 0, pending: false };
  try {
    const parsed = JSON.parse(
      safeGetLocalStorage(localMetaKey(key)) || "{}",
    );
    return {
      editedAt: Number(parsed.editedAt || 0),
      pending: Boolean(parsed.pending),
    };
  } catch {
    return { editedAt: 0, pending: false };
  }
}

function setLocalMeta(
  key: string,
  meta: { editedAt: number; pending: boolean },
) {
  if (typeof window === "undefined") return;
  safeSetLocalStorage(localMetaKey(key), JSON.stringify(meta));
}

async function saveSharedState<T>(key: string, value: T) {
  const config = sharedApiConfig();
  const response = await fetch(`${config.endpoint}/${encodeURIComponent(key)}`, {
    method: "PUT",
    headers: config.headers,
    body: JSON.stringify({ data: value }),
  });

  if (!response.ok) throw new Error(`D1 save failed: ${response.status}`);
  const row = (await response.json()) as AppStateRow<T>;
  return row.updated_at || new Date().toISOString();
}

function reportSharedSaveError(error: unknown) {
  // 수정·자동저장 과정에서 팝업이 반복되지 않도록 콘솔 기록만 남깁니다.
  console.warn("공유 저장소(Cloudflare D1) 저장 실패", error);
}

function useLocal<T>(key: string, initial: T) {
  const [value, rawSetValue] = useState<T>(initial);
  const [loaded, setLoaded] = useState(false);
  const valueRef = useRef<T>(initial);
  const saveTimerRef = useRef<number | null>(null);
  const isSavingRef = useRef(false);
  const keyRef = useRef(key);
  const lastRemoteUpdatedAtRef = useRef<string | null>(null);
  const isReloadingRemoteRef = useRef(false);
  const alertOpenRef = useRef(false);
  const lastRemoteCheckAtRef = useRef(0);

  useEffect(() => {
    keyRef.current = key;
  }, [key]);

  const applyValue = (nextValue: T, cacheLocal = true) => {
    valueRef.current = nextValue;
    rawSetValue(nextValue);
    if (cacheLocal && typeof window !== "undefined") {
      safeSetLocalStorage(keyRef.current, JSON.stringify(nextValue));
    }
  };

  const refreshFromSupabase = async (options?: { allowDuringPending?: boolean }) => {
    if (typeof window === "undefined") return;

    const meta = getLocalMeta(keyRef.current);
    if (meta.pending && !options?.allowDuringPending) return;
    if (isSavingRef.current && !options?.allowDuringPending) return;

    isReloadingRemoteRef.current = true;
    try {
      const remoteRow = await loadSharedStateRow<T>(keyRef.current);
      if (!remoteRow || remoteRow.data === null || remoteRow.data === undefined) return;
      lastRemoteUpdatedAtRef.current = remoteRow.updated_at || null;
      applyValue(remoteRow.data);
      setLocalMeta(keyRef.current, {
        editedAt: Date.now(),
        pending: false,
      });
    } finally {
      isReloadingRemoteRef.current = false;
    }
  };

  const persistNow = (nextValue: T) => {
    if (typeof window === "undefined") return;

    const editedAt = Date.now();
    valueRef.current = nextValue;
    safeSetLocalStorage(keyRef.current, JSON.stringify(nextValue));
    setLocalMeta(keyRef.current, { editedAt, pending: true });

    if (saveTimerRef.current) window.clearTimeout(saveTimerRef.current);
    saveTimerRef.current = window.setTimeout(async () => {
      try {
        isSavingRef.current = true;
        const savedUpdatedAt = await saveSharedState(keyRef.current, nextValue);
        if (savedUpdatedAt) lastRemoteUpdatedAtRef.current = savedUpdatedAt;
        const meta = getLocalMeta(keyRef.current);
        if (meta.editedAt === editedAt)
          setLocalMeta(keyRef.current, { editedAt, pending: false });
      } catch (error) {
        const meta = getLocalMeta(keyRef.current);
        if (meta.editedAt === editedAt)
          setLocalMeta(keyRef.current, { editedAt, pending: true });
        reportSharedSaveError(error);
      } finally {
        isSavingRef.current = false;
      }
    }, 180);
  };

  const setValue: React.Dispatch<React.SetStateAction<T>> = (next) => {
    rawSetValue((prev) => {
      const resolved =
        typeof next === "function" ? (next as (prev: T) => T)(prev) : next;
      valueRef.current = resolved;
      persistNow(resolved);
      return resolved;
    });
  };

  const checkRemoteUpdate = async () => {
    if (typeof window === "undefined") return;
    if (alertOpenRef.current || isSavingRef.current || isReloadingRemoteRef.current) return;

    const meta = getLocalMeta(keyRef.current);
    if (meta.pending) return;

    lastRemoteCheckAtRef.current = Date.now();
    const remoteRow = await loadSharedStateRow<T>(keyRef.current);
    if (!remoteRow?.updated_at) return;

    if (!lastRemoteUpdatedAtRef.current) {
      lastRemoteUpdatedAtRef.current = remoteRow.updated_at;
      return;
    }

    if (remoteRow.updated_at === lastRemoteUpdatedAtRef.current) return;

    alertOpenRef.current = true;
    try {
      // 로컬에 저장 대기 중인 수정이 없을 때만 최신 원격값을 조용히 반영합니다.
      lastRemoteUpdatedAtRef.current = remoteRow.updated_at;
      applyValue(remoteRow.data);
      setLocalMeta(keyRef.current, { editedAt: Date.now(), pending: false });
    } finally {
      alertOpenRef.current = false;
    }
  };

  useEffect(() => {
    let cancelled = false;

    async function hydrate() {
      try {
        // Cloudflare D1을 원본으로 사용합니다. 로컬 캐시가 있어도 원격 최신값을 먼저 불러옵니다.
        const remoteRow = await loadSharedStateRow<T>(key);
        if (!cancelled && remoteRow?.data !== null && remoteRow?.data !== undefined) {
          keyRef.current = key;
          lastRemoteUpdatedAtRef.current = remoteRow.updated_at || null;
          applyValue(remoteRow.data);
          setLocalMeta(key, { editedAt: Date.now(), pending: false });
          return;
        }

        // Cloudflare D1에 아직 데이터가 없을 때만 로컬 캐시를 임시로 사용합니다.
        const localSaved = safeGetLocalStorage(key);
        if (localSaved && !cancelled) {
          const parsed = JSON.parse(localSaved) as T;
          keyRef.current = key;
          applyValue(parsed, false);
          return;
        }
      } catch (error) {
        console.warn(
          "공유 데이터 불러오기 실패, 브라우저 저장소를 임시로 사용합니다.",
          error,
        );
        const localSaved = safeGetLocalStorage(key);
        if (localSaved && !cancelled) {
          try {
            const parsed = JSON.parse(localSaved) as T;
            keyRef.current = key;
            applyValue(parsed, false);
          } catch {
            // 로컬 캐시도 손상된 경우 초기값을 유지합니다.
          }
        }
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
    if (!loaded || typeof window === "undefined") return;

    const retryPendingSave = async () => {
      try {
        const meta = getLocalMeta(key);
        if (!meta.pending) return;
        const localSaved = safeGetLocalStorage(key);
        if (!localSaved) return;
        const parsed = JSON.parse(localSaved) as T;
        const savedUpdatedAt = await saveSharedState(key, parsed);
        if (savedUpdatedAt) lastRemoteUpdatedAtRef.current = savedUpdatedAt;
        setLocalMeta(key, { editedAt: meta.editedAt, pending: false });
      } catch (error) {
        console.warn("공유 데이터 저장 재시도 실패", error);
      }
    };

    const safeCheckRemoteUpdate = () => {
      checkRemoteUpdate().catch((error) =>
        console.warn("공유 데이터 변경 확인 실패", error),
      );
    };

    const refreshIfStale = () => {
      if (Date.now() - lastRemoteCheckAtRef.current < FOCUS_REFRESH_MIN_GAP_MS) return;
      safeCheckRemoteUpdate();
    };

    const onVisibilityChange = () => {
      if (document.visibilityState === "visible") refreshIfStale();
    };

    const interval = window.setInterval(() => {
      retryPendingSave();
      safeCheckRemoteUpdate();
    }, STATIC_SYNC_INTERVAL_MS);
    window.addEventListener("focus", refreshIfStale);
    document.addEventListener("visibilitychange", onVisibilityChange);

    return () => {
      window.clearInterval(interval);
      window.removeEventListener("focus", refreshIfStale);
      document.removeEventListener("visibilitychange", onVisibilityChange);
    };
  }, [key, loaded]);

  return [value, setValue] as const;
}



function salesChunkKey(baseKey: string, period: PeriodType, month: string) {
  return `${baseKey}_${period}_${month}`;
}

function chunkMonthForSale(row: SalesRecord) {
  if (row.period === "current") return row.refMonth || row.saleDate.slice(0, 7);
  return row.refMonth || row.saleDate.slice(0, 7);
}

function salesStorageKey(baseKey: string, row: SalesRecord) {
  return salesChunkKey(baseKey, row.period, chunkMonthForSale(row));
}

function dedupeSales(records: SalesRecord[]) {
  const map = new Map<string, SalesRecord>();
  records.forEach((row) => map.set(row.id, row));
  return Array.from(map.values());
}

function salesKeysForMonth(baseKey: string, month: string) {
  return [
    salesChunkKey(baseKey, "current", month),
    salesChunkKey(baseKey, "current", addMonths(month, 1)),
    salesChunkKey(baseKey, "prevMonth", month),
    salesChunkKey(baseKey, "prevYear", month),
  ];
}

function groupSalesByStorageKey(baseKey: string, records: SalesRecord[]) {
  const map = new Map<string, SalesRecord[]>();
  records.forEach((row) => {
    const key = salesStorageKey(baseKey, row);
    map.set(key, [...(map.get(key) || []), row]);
  });
  return map;
}

function shouldShowSalesForMonth(row: SalesRecord, month: string) {
  if (row.period === "current") {
    const rowMonth = row.refMonth || row.saleDate.slice(0, 7);
    return rowMonth === month || rowMonth === addMonths(month, 1);
  }
  return row.refMonth === month;
}

type V3SalesRecordRow = {
  row_key: string;
  period: PeriodType;
  ref_month: string;
  sale_date: string;
  store_code: string;
  store_name: string;
  channel: string;
  manager: string;
  store_type: string;
  brand: string;
  item_code: string;
  item_name: string;
  quantity: number | string;
  sales_amount: number | string;
  cost_amount: number | string;
  profit_amount: number | string;
  profit_rate: number | string;
};

function d1SalesEndpoint(path: string) {
  return {
    url: `/api/d1/${path.replace(/^\/+/, "")}`,
    headers: { "Content-Type": "application/json" },
  };
}

function toV3Payload(row: SalesRecord) {
  return {
    id: row.id,
    period: row.period,
    ref_month: row.refMonth,
    sale_date: row.saleDate,
    store_code: row.storeCode,
    store_name: row.storeName,
    channel: row.channel,
    manager: row.manager,
    store_type: row.storeType,
    brand: row.brand,
    item_code: row.itemCode,
    item_name: row.itemName,
    quantity: Number(row.quantity || 0),
    sales_amount: Number(row.salesAmount || 0),
    cost_amount: Number(row.costAmount || 0),
    profit_amount: Number(row.profitAmount || 0),
    profit_rate: Number(row.profitRate || 0),
  };
}

function fromV3Row(row: V3SalesRecordRow): SalesRecord {
  return {
    id: row.row_key,
    period: row.period,
    refMonth: row.ref_month,
    saleDate: row.sale_date,
    storeCode: row.store_code,
    storeName: row.store_name,
    channel: row.channel,
    manager: row.manager,
    storeType: row.store_type,
    brand: row.brand,
    itemCode: row.item_code,
    itemName: row.item_name,
    quantity: Number(row.quantity || 0),
    salesAmount: Number(row.sales_amount || 0),
    costAmount: Number(row.cost_amount || 0),
    profitAmount: Number(row.profit_amount || 0),
    profitRate: Number(row.profit_rate || 0),
  };
}

async function loadV3SalesForMonth(month: string): Promise<{
  available: boolean;
  periodsWithBatch: Set<PeriodType>;
  records: SalesRecord[];
}> {
  const api = d1SalesEndpoint("sales");
  const response = await fetch(
    `${api.url}?baseMonth=${encodeURIComponent(month)}`,
    { method: "GET", headers: api.headers, cache: "no-store" },
  );
  if (response.status === 404) {
    return { available: false, periodsWithBatch: new Set(), records: [] };
  }
  if (!response.ok) throw new Error(`D1 sales load failed: ${response.status}`);
  const payload = (await response.json()) as {
    available?: boolean;
    records?: V3SalesRecordRow[];
    batches?: { period: PeriodType }[];
  };
  return {
    available: payload.available !== false,
    periodsWithBatch: new Set((payload.batches || []).map((row) => row.period)),
    records: (payload.records || []).map(fromV3Row),
  };
}

async function replaceV3SalesBatch(request: SalesUploadRequest) {
  const api = d1SalesEndpoint("sales/replace");
  const response = await fetch(api.url, {
    method: "POST",
    headers: api.headers,
    body: JSON.stringify({
      period: request.period,
      refMonth: request.refMonth,
      fileName: request.fileName,
      uploadedDates: request.uploadedDates,
      rows: request.rows,
    }),
  });
  if (!response.ok) {
    const body = await response.text();
    throw new Error(`D1 sales upload failed: ${response.status} ${body.slice(0, 300)}`);
  }
  return true;
}

async function deleteV3CurrentDate(refMonth: string, saleDate: string) {
  const api = d1SalesEndpoint("sales/delete-date");
  const response = await fetch(api.url, {
    method: "POST",
    headers: api.headers,
    body: JSON.stringify({ refMonth, saleDate }),
  });
  if (!response.ok) {
    const body = await response.text();
    throw new Error(`D1 sales delete failed: ${response.status} ${body.slice(0, 300)}`);
  }
  return true;
}

function useChunkedSales(
  baseKey: string,
  initial: SalesRecord[],
  month: string,
) {
  const [value, rawSetValue] = useState<SalesRecord[]>(initial);
  const [storageMode, setStorageMode] = useState<"checking" | "v3" | "legacy">("checking");
  const valueRef = useRef<SalesRecord[]>(initial);
  const saveTimerRef = useRef<number | null>(null);
  const isSavingRef = useRef(false);
  const monthRef = useRef(month);
  const lastSalesRefreshAtRef = useRef(0);

  useEffect(() => {
    monthRef.current = month;
  }, [month]);

  const loadLegacyMonthChunks = async (targetMonth: string) => {
    const keys = salesKeysForMonth(baseKey, targetMonth);
    const rows = await Promise.all(keys.map((key) => loadSharedStateRow<SalesRecord[]>(key)));
    const chunkedSales = rows.filter(Boolean).flatMap((row) => row?.data || []);
    if (chunkedSales.length) return dedupeSales(chunkedSales);
    const legacy = await loadSharedState<SalesRecord[]>(baseKey);
    if (legacy?.length)
      return dedupeSales(legacy.filter((row) => shouldShowSalesForMonth(row, targetMonth)));
    return initial.filter((row) => shouldShowSalesForMonth(row, targetMonth));
  };

  const loadMonthData = async (targetMonth: string) => {
    const v3 = await loadV3SalesForMonth(targetMonth);
    if (!v3.available) {
      setStorageMode("legacy");
      return loadLegacyMonthChunks(targetMonth);
    }

    setStorageMode("v3");
    const legacy = await loadLegacyMonthChunks(targetMonth);
    const v3Periods = v3.periodsWithBatch;
    const legacyForUnmigratedPeriods = legacy.filter((row) => !v3Periods.has(row.period));
    return dedupeSales([...legacyForUnmigratedPeriods, ...v3.records]);
  };

  const persistLegacyNow = (nextValue: SalesRecord[]) => {
    if (typeof window === "undefined") return;
    const editedAt = Date.now();
    valueRef.current = nextValue;
    setLocalMeta(`${baseKey}_chunked`, { editedAt, pending: true });
    if (saveTimerRef.current) window.clearTimeout(saveTimerRef.current);
    saveTimerRef.current = window.setTimeout(async () => {
      try {
        isSavingRef.current = true;
        const grouped = groupSalesByStorageKey(baseKey, nextValue);
        const targetKeys = salesKeysForMonth(baseKey, monthRef.current);
        await Promise.all(
          targetKeys.map((key) => saveSharedState(key, grouped.get(key) || [])),
        );
        const meta = getLocalMeta(`${baseKey}_chunked`);
        if (meta.editedAt === editedAt)
          setLocalMeta(`${baseKey}_chunked`, { editedAt, pending: false });
      } catch (error) {
        setLocalMeta(`${baseKey}_chunked`, { editedAt, pending: true });
        reportSharedSaveError(error);
      } finally {
        isSavingRef.current = false;
      }
    }, 180);
  };

  const setValue: React.Dispatch<React.SetStateAction<SalesRecord[]>> = (next) => {
    rawSetValue((prev) => {
      const resolved = typeof next === "function"
        ? (next as (prev: SalesRecord[]) => SalesRecord[])(prev)
        : next;
      const deduped = dedupeSales(resolved);
      valueRef.current = deduped;
      // 일반 화면 편집은 기존 방식과 호환합니다. 매출 업로드/삭제는 actions의 원자적 RPC를 사용합니다.
      persistLegacyNow(deduped);
      return deduped;
    });
  };

  const refreshFromSupabase = async (targetMonth = monthRef.current) => {
    if (isSavingRef.current) return;
    lastSalesRefreshAtRef.current = Date.now();
    try {
      const latest = await loadMonthData(targetMonth);
      valueRef.current = latest;
      rawSetValue(latest);
      setLocalMeta(`${baseKey}_chunked`, { editedAt: Date.now(), pending: false });
    } catch (error) {
      console.warn("매출 데이터 동기화 실패", error);
    }
  };

  const replaceUpload: SalesStorageActions["replaceUpload"] = async (request) => {
    isSavingRef.current = true;
    try {
      const currentValue = valueRef.current;
      const next = request.period === "current"
        ? currentValue.filter((row) => !(row.period === "current" && request.uploadedDates.includes(row.saleDate)))
        : currentValue.filter((row) => !(row.period === request.period && row.refMonth === request.refMonth));
      const merged = dedupeSales([...next, ...request.rows]);
      const savedToV3 = await replaceV3SalesBatch(request);
      if (savedToV3) {
        setStorageMode("v3");
        valueRef.current = merged;
        rawSetValue(merged);
        return { mode: "v3" };
      }
      setStorageMode("legacy");
      valueRef.current = merged;
      rawSetValue(merged);
      persistLegacyNow(merged);
      return { mode: "legacy" };
    } finally {
      isSavingRef.current = false;
    }
  };

  const deleteCurrentDate: SalesStorageActions["deleteCurrentDate"] = async (refMonth, saleDate) => {
    isSavingRef.current = true;
    try {
      const next = valueRef.current.filter(
        (row) => !(row.period === "current" && row.refMonth === refMonth && row.saleDate === saleDate),
      );
      const deletedInV3 = await deleteV3CurrentDate(refMonth, saleDate);
      if (deletedInV3) {
        setStorageMode("v3");
        valueRef.current = next;
        rawSetValue(next);
        return { mode: "v3" };
      }
      setStorageMode("legacy");
      valueRef.current = next;
      rawSetValue(next);
      persistLegacyNow(next);
      return { mode: "legacy" };
    } finally {
      isSavingRef.current = false;
    }
  };

  useEffect(() => {
    let cancelled = false;
    async function hydrateSales() {
      try {
        const latest = await loadMonthData(month);
        if (cancelled) return;
        valueRef.current = latest;
        rawSetValue(latest);
        lastSalesRefreshAtRef.current = Date.now();
      } catch (error) {
        console.warn("매출 데이터 불러오기 실패", error);
        setStorageMode("legacy");
      }
    }
    hydrateSales();
    return () => {
      cancelled = true;
      if (saveTimerRef.current) window.clearTimeout(saveTimerRef.current);
    };
  }, [baseKey, month]);

  useEffect(() => {
    if (typeof window === "undefined") return;
    const interval = window.setInterval(() => {
      refreshFromSupabase().catch((error) => console.warn("매출 데이터 자동 동기화 실패", error));
    }, SALES_SYNC_INTERVAL_MS);
    const onFocus = () => {
      if (Date.now() - lastSalesRefreshAtRef.current < FOCUS_REFRESH_MIN_GAP_MS) return;
      refreshFromSupabase();
    };
    window.addEventListener("focus", onFocus);
    return () => {
      window.clearInterval(interval);
      window.removeEventListener("focus", onFocus);
    };
  }, [baseKey]);

  const actions: SalesStorageActions = {
    replaceUpload,
    deleteCurrentDate,
    refresh: () => refreshFromSupabase(),
    storageMode,
  };

  return [value, setValue, actions] as const;
}

function readFileRows(file: File): Promise<Record<string, unknown>[]> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: "array", cellDates: true });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
          header: 1,
          defval: "",
        });
        const headerIndex = matrix.findIndex((row) =>
          row.some((cell) =>
            [
              "거래처코드",
              "거래처명",
              "주문일",
              "판매일",
              "매출일",
              "Target",
              "EST",
            ].includes(norm(cell)),
          ),
        );

        if (headerIndex < 0) {
          resolve(
            XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, {
              defval: "",
            }),
          );
          return;
        }

        const headers = matrix[headerIndex].map((h) => norm(h));
        const rows = matrix
          .slice(headerIndex + 1)
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

function sum(
  records: SalesRecord[],
  key: keyof Pick<
    SalesRecord,
    "salesAmount" | "costAmount" | "profitAmount" | "quantity"
  >,
) {
  return records.reduce((a, b) => a + Number(b[key] || 0), 0);
}

function profitRateValue(value: unknown) {
  if (value === undefined || value === null || norm(value) === "")
    return undefined;

  if (typeof value === "number") {
    return Math.abs(value) > 0 && Math.abs(value) <= 1 ? value * 100 : value;
  }

  const text = norm(value);
  const parsed = num(text);
  return text.includes("%")
    ? parsed
    : Math.abs(parsed) > 0 && Math.abs(parsed) <= 1
      ? parsed * 100
      : parsed;
}

function weightedProfitRate(records: SalesRecord[]) {
  const salesTotal = sum(records, "salesAmount");
  if (!salesTotal) return 0;

  const weighted = records.reduce(
    (total, row) =>
      total + Number(row.salesAmount || 0) * Number(row.profitRate || 0),
    0,
  );
  return weighted / salesTotal;
}

function exportExcel(
  rows: Record<string, string | number>[],
  fileName: string,
) {
  if (!rows.length) {
    alert("다운로드할 데이터가 없습니다.");
    return;
  }

  const ws = XLSX.utils.json_to_sheet(rows);
  ws["!cols"] = Object.keys(rows[0]).map((key) => ({
    wch: Math.max(
      12,
      key.length + 4,
      ...rows.map((row) => String(row[key] ?? "").length + 2),
    ),
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
    기간구분:
      r.period === "current"
        ? "당월"
        : r.period === "prevMonth"
          ? "전월"
          : "전년동월",
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
const EST_ENTRY_MANAGERS: Manager[] = ["SY", "KT", "NH"];

function isEstEntryPeriodToday() {
  return Number(today().slice(8, 10)) <= 4;
}

export default function SalesReportClient() {
  const [active, setActive] = useState("대시보드");
  const [isAdmin, setIsAdmin] = useState(false);
  const [dashMonth, setDashMonth] = useState(thisMonth());
  const [dashDate, setDashDate] = useState(today());
  const [stores, setStores] = useLocal<Store[]>(
    "ablab_stores_v15",
    initialStores,
  );
  const [sales, setSales, salesActions] = useChunkedSales(
    "ablab_sales_v14",
    initialSales,
    dashMonth,
  );
  const [targets, setTargets] = useLocal<TargetRecord[]>(
    "ablab_targets_v14",
    initialTargets,
  );
  const [ests, setEsts] = useLocal<EstRecord[]>("ablab_ests_v14", initialEsts);
  const [timeConfigs, setTimeConfigs] = useLocal<TimeConfig[]>(
    "ablab_time_configs_v14",
    initialTimeConfigs,
  );
  const [codeMappings, setCodeMappings] = useLocal<StoreCodeMapping[]>(
    "ablab_code_mappings_v1",
    [],
  );
  const [itemCosts, setItemCosts] = useLocal<ItemCostRecord[]>(
    "ablab_item_costs_v1",
    initialItemCosts,
  );
  const [itemMasters, setItemMasters] = useLocal<ItemMasterRecord[]>(
    "ablab_item_masters_v1",
    [],
  );

  useEffect(() => {
    if (!dashDate.startsWith(dashMonth)) setDashDate(monthEnd(dashMonth));
  }, [dashMonth, dashDate]);

  const isEstEntryOpen = isEstEntryPeriodToday();
  const canAccessEstEntry = isAdmin || isEstEntryOpen;

  useEffect(() => {
    if (!isAdmin && active === "월초관리") setActive("대시보드");
    if (!canAccessEstEntry && active === "EST 입력") setActive("대시보드");
  }, [isAdmin, active, canAccessEstEntry]);

  const tg = useMemo(
    () => getTimeGone(dashMonth, dashDate, timeConfigs),
    [dashMonth, dashDate, timeConfigs],
  );
  const menus = [
    ...(canAccessEstEntry ? [{ label: "EST 입력", order: "0" }] : []),
    { label: "대시보드", order: "1" },
    { label: "매출현황", order: "2" },
    { label: "거래처별 상세", order: "3" },
    { label: "품목분석", order: "4" },
    { label: "매입가 정보", order: "5" },
    ...(isAdmin ? [{ label: "월초관리", order: "6" }] : []),
  ];

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
    <main
      className="sales-report-root h-screen overflow-hidden bg-white text-black"
      data-active={active}
      style={{ fontFamily: '"Malgun Gothic", "맑은 고딕", sans-serif' }}
    >
      <header className="sticky top-0 z-40 border-b border-gray-200 bg-white shadow-sm">
        <div className="flex flex-col gap-3 px-4 py-3 lg:flex-row lg:items-center lg:justify-between lg:px-5">
          <div className="flex flex-wrap items-center gap-3">
            <div className="text-lg font-extrabold tracking-tight text-orange-950">
              에이비랩 코리아 Sales Report
            </div>
            <nav className="flex flex-wrap items-center gap-2">
              {menus.map((m) => (
                <button
                  key={m.label}
                  onClick={() => setActive(m.label)}
                  className={`rounded-xl px-4 py-2 text-xs font-bold transition ${
                    active === m.label
                      ? "bg-orange-500 text-white shadow"
                      : "bg-orange-50 text-orange-900 hover:bg-orange-100"
                  }`}
                >
                  {m.order}. {m.label}
                </button>
              ))}
            </nav>
          </div>
          <div>
            {isAdmin ? (
              <button
                type="button"
                onClick={() => {
                  setIsAdmin(false);
                  setActive("대시보드");
                }}
                className="rounded-xl bg-orange-100 px-4 py-2 text-xs font-bold text-orange-900 hover:bg-orange-200"
              >
                관리자 모드 해제
              </button>
            ) : (
              <button
                type="button"
                onClick={adminLogin}
                className="rounded-xl bg-orange-100 px-4 py-2 text-xs font-bold text-orange-900 hover:bg-orange-200"
              >
                관리자 로그인
              </button>
            )}
          </div>
        </div>
      </header>

      <section className="flex h-[calc(100vh-69px)] min-w-0 flex-col overflow-hidden p-4 lg:p-5">
        <style jsx global>{`
          .sales-report-root table th,
          .sales-report-root table td,
          .sales-report-root table input,
          .sales-report-root table select,
          .sales-report-root .metric-black { color: #000 !important; }
          .sales-report-root table th,
          .sales-report-root table td { border-color: #e5e7eb !important; }

          /* 임원·전사 공유용 표 헤더: 색은 분류만 돕고 숫자보다 먼저 보이지 않도록 절제합니다. */
          .sales-report-root[data-active="대시보드"] table thead th,
          .sales-report-root[data-active="매출현황"] table thead th,
          .sales-report-root[data-active="거래처별 상세"] table thead th,
          .sales-report-root[data-active="품목분석"] table thead th {
            border-width: 1px !important;
            border-style: solid !important;
            border-color: #e5e7eb !important;
            border-radius: 0 !important;
            color: #111827 !important;
            font-weight: 700;
            background-clip: border-box !important;
            box-shadow: inset 0 -1px 0 rgba(148, 163, 184, 0.18) !important;
          }

          /* 표 섹션 제목은 색 대신 간격과 타이포그래피로 위계를 만듭니다. */
          .sales-report-root[data-active="대시보드"] h2,
          .sales-report-root[data-active="대시보드"] h3,
          .sales-report-root[data-active="매출현황"] h2,
          .sales-report-root[data-active="매출현황"] h3,
          .sales-report-root[data-active="거래처별 상세"] h2,
          .sales-report-root[data-active="거래처별 상세"] h3,
          .sales-report-root[data-active="품목분석"] h2,
          .sales-report-root[data-active="품목분석"] h3 {
            color: #111827;
            letter-spacing: -0.01em;
          }

          /* 모든 카테고리 공통: 표 내부 스크롤 시 고정 헤더 뒤로 본문이 비치지 않도록 처리 */
          .sales-report-root .overflow-auto:has(table) {
            position: relative;
            isolation: isolate;
            background: #fff;
            overscroll-behavior: contain;
            scrollbar-gutter: stable;
          }
          .sales-report-root table {
            position: relative;
            background: #fff;
          }
          .sales-report-root table thead {
            position: relative;
            z-index: 70;
          }
          .sales-report-root table thead tr {
            position: relative;
            z-index: 70;
          }
          .sales-report-root table thead th {
            opacity: 1 !important;
            background-clip: border-box !important;
            isolation: isolate;
            box-shadow: inset 0 -1px 0 #e5e7eb, 0 1px 2px rgba(15, 23, 42, 0.04);
          }
          .sales-report-root table thead th::before {
            content: "";
            position: absolute;
            inset: 0;
            z-index: -1;
            background: inherit;
            pointer-events: none;
          }
          .sales-report-root table thead th::after {
            content: "";
            position: absolute;
            left: -1px;
            right: -1px;
            bottom: -1px;
            height: 2px;
            z-index: 2;
            background: #e5e7eb;
            pointer-events: none;
          }
          .sales-report-root table tbody {
            position: relative;
            z-index: 1;
          }
          .sales-report-root table tr.sticky > th,
          .sales-report-root table tr.sticky > td {
            opacity: 1 !important;
            background-clip: border-box !important;
            box-shadow: inset 0 -1px 0 #e5e7eb, 0 1px 2px rgba(15, 23, 42, 0.04);
          }
          .sales-report-root table tr.sticky > th::before,
          .sales-report-root table tr.sticky > td::before {
            content: "";
            position: absolute;
            inset: 0;
            z-index: -1;
            background: inherit;
            pointer-events: none;
          }
          /* 매출현황 2단 헤더: 품목분석과 동일하게 상단 그룹과 하단 세부 헤더 사이 구분선 표시 */
          .sales-report-root .sales-status-table thead tr:first-child th {
            height: 31px;
            min-height: 31px;
            padding-top: 0 !important;
            padding-bottom: 0 !important;
            box-shadow: none !important;
          }
          .sales-report-root .sales-status-table thead tr:first-child th:not([rowspan]) {
            border-bottom: 1px solid #d7dee7 !important;
            box-shadow: inset 0 -1px 0 #d7dee7 !important;
          }
          .sales-report-root .sales-status-table thead tr:first-child th::after {
            display: none !important;
          }
          .sales-report-root .sales-status-table thead tr:nth-child(2) th {
            top: 31px !important;
            border-top: 1px solid #cbd5e1 !important;
            box-shadow: 0 1px 2px rgba(15, 23, 42, 0.06) !important;
          }
          .sales-report-root .sales-status-table thead tr:nth-child(2) th::before {
            top: 0;
          }
          .sales-report-root .sales-status-table thead tr:nth-child(2) th::after {
            bottom: -1px;
          }

          /* 거래처별 상세 2단 헤더: 품목분석과 동일하게 상단 그룹과 하단 세부 헤더 사이 구분선 표시 */
          .sales-report-root .connected-two-tier thead tr:first-child th {
            height: 37px;
            min-height: 37px;
            box-shadow: none !important;
            text-align: center !important;
          }
          .sales-report-root .connected-two-tier thead tr:first-child th:not([rowspan]) {
            border-bottom: 1px solid #d7dee7 !important;
            box-shadow: inset 0 -1px 0 #d7dee7 !important;
          }
          .sales-report-root .connected-two-tier thead tr:first-child th::after {
            display: none !important;
          }
          .sales-report-root .connected-two-tier thead tr:nth-child(2) th {
            top: 37px !important;
            border-top: 1px solid #cbd5e1 !important;
            text-align: center !important;
            box-shadow: 0 1px 2px rgba(15, 23, 42, 0.06) !important;
          }
          .sales-report-root .connected-two-tier thead tr:nth-child(2) th::before {
            top: 0;
          }
          .sales-report-root .connected-two-tier thead th,
          .sales-report-root .connected-two-tier thead button {
            text-align: center !important;
            justify-content: center !important;
          }
          /* 2단 헤더 기간 그룹 구분선: sticky 위치는 유지하고 그룹 경계만 또렷하게 표시 */
          .sales-report-root .period-group-start {
            border-left: 2px solid #cbd5e1 !important;
          }
          .sales-report-root .period-subgroup-start {
            border-left: 2px solid #d7dee8 !important;
          }
          /* 품목분석 손익요약: 헤더·SUBTOTAL과 본문 스크롤 영역을 완전히 분리 */
          .sales-report-root .item-profit-fixed-header,
          .sales-report-root .item-profit-fixed-body {
            table-layout: fixed;
            border-collapse: collapse;
            border-spacing: 0;
            font-size: 12px;
          }
          .sales-report-root .item-profit-fixed-header-wrap {
            padding-right: 17px;
            background: #fff;
          }
          .sales-report-root .item-profit-fixed-header {
            position: relative;
            z-index: 30;
            background: #fff;
          }
          .sales-report-root .item-profit-fixed-header thead tr:first-child {
            height: 34px;
          }
          .sales-report-root .item-profit-fixed-header thead tr:nth-child(2) {
            height: 44px;
          }
          .sales-report-root .item-profit-fixed-header thead tr:nth-child(3) {
            height: 34px;
          }
          .sales-report-root .item-profit-fixed-header thead,
          .sales-report-root .item-profit-fixed-header thead tr {
            position: static !important;
            top: auto !important;
          }
          .sales-report-root .item-profit-fixed-header thead th {
            position: relative !important;
            top: auto !important;
            z-index: 1 !important;
            overflow: hidden;
            background-clip: border-box !important;
          }
          /* 공통 sticky 헤더용 가상 배경이 분리형 헤더 전체를 덮지 않도록 차단 */
          .sales-report-root .item-profit-fixed-header thead th::before,
          .sales-report-root .item-profit-fixed-header thead th::after,
          .sales-report-root .item-profit-fixed-header .item-profit-subtotal th::before,
          .sales-report-root .item-profit-fixed-header .item-profit-subtotal th::after {
            display: none !important;
            content: none !important;
          }
          .sales-report-root .item-profit-fixed-header .item-profit-subtotal th {
            position: relative !important;
            top: auto !important;
            z-index: 2 !important;
            background: #fff8dc !important;
            border: 1px solid #e5e7eb !important;
            border-radius: 7px;
            box-shadow: 0 1px 2px rgba(15, 23, 42, 0.06);
          }
          .sales-report-root .item-profit-fixed-header .item-profit-subtotal .subtotal-label {
            padding-left: 16px !important;
            text-align: left !important;
            letter-spacing: 0.01em;
          }
          .sales-report-root .item-profit-fixed-header .item-profit-subtotal .subtotal-number {
            padding-right: 12px !important;
            text-align: right !important;
          }
          .sales-report-root .item-profit-fixed-body .item-profit-number-cell {
            padding-right: 12px !important;
            text-align: right !important;
          }
          .sales-report-root .item-profit-fixed-body tbody {
            position: static !important;
          }
          .sales-report-root .item-profit-fixed-body td {
            height: 40px;
            padding: 8px;
            font-size: 12px;
            vertical-align: middle;
          }
          .sales-report-root .item-profit-fixed-header th,
          .sales-report-root .item-profit-fixed-body td {
            box-sizing: border-box;
          }
          .sales-report-root .item-profit-fixed-body .item-name-cell {
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
          }

          /* 대시보드 EST 카드: 담당자별 매출요약의 EST 계열과 동일한 색상 */
          .sales-report-root .est-summary-card {
            border-color: #f2d675 !important;
            background: #fffdf2 !important;
            box-shadow: 0 1px 2px rgba(15, 23, 42, 0.06) !important;
          }

        `}</style>
        <div
          className={
            ["매출현황", "거래처별 상세", "품목분석", "매입가 정보"].includes(active)
              ? "mb-2 space-y-1"
              : "mb-4 space-y-3 rounded-2xl border border-gray-200 bg-white p-4 shadow-sm"
          }
        >
          <div
            className={
              ["매출현황", "거래처별 상세", "품목분석", "매입가 정보"].includes(active)
                ? "bg-white px-0 py-0"
                : "rounded-2xl border border-gray-200 bg-slate-50 p-3 shadow-sm"
            }
          >
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
              <HeaderMetricInfo
                title="TIME GONE"
                value={pct(tg.timeGoneRate)}
              />
              <HeaderTimeInfo title="총일수" value={tg.totalDays} />
              <HeaderTimeInfo title="진행일수" value={tg.progressedDays} />
              <HeaderTimeInfo title="잔여일수" value={tg.remainingDays} />
              {active === "매출현황" && (
                <SalesTargetHeaderKpi
                  stores={stores}
                  sales={sales}
                  targets={targets}
                  ests={ests}
                  month={dashMonth}
                  date={dashDate}
                />
              )}
            </div>
          </div>

          {active === "대시보드" && (
            <DashboardTopKpis
              stores={stores}
              sales={sales}
              targets={targets}
              ests={ests}
              itemCosts={itemCosts}
              month={dashMonth}
              date={dashDate}
            />
          )}
        </div>

        {active === "EST 입력" && (
          <EstQuickEntry
            stores={stores}
            sales={sales}
            ests={ests}
            setEsts={setEsts}
            targets={targets}
            setTargets={setTargets}
            month={dashMonth}
            canEdit={isAdmin || isEstEntryOpen}
            isAdmin={isAdmin}
          />
        )}
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
        {active === "거래처별 상세" && (
          <ItemAnalysis
            stores={stores}
            sales={sales}
            month={dashMonth}
            date={dashDate}
            pageTitle="거래처별 상세"
          />
        )}
        {active === "품목분석" && (
          <ItemShipmentAnalysis
            stores={stores}
            sales={sales}
            itemMasters={itemMasters}
            month={dashMonth}
            date={dashDate}
          />
        )}
        {active === "매입가 정보" && (
          <ItemCostStatus
            sales={sales}
            itemCosts={itemCosts}
            setItemCosts={setItemCosts}
            isAdmin={isAdmin}
          />
        )}
        {isAdmin && active === "월초관리" && (
          <MonthStartManagement
            stores={stores}
            setStores={setStores}
            sales={sales}
            setSales={setSales}
            salesActions={salesActions}
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
            itemMasters={itemMasters}
            setItemMasters={setItemMasters}
          />
        )}
      </section>
    </main>
  );
}


function EstQuickEntry({
  stores,
  sales,
  ests,
  setEsts,
  targets,
  setTargets,
  month,
  canEdit,
  isAdmin,
}: {
  stores: Store[];
  sales: SalesRecord[];
  ests: EstRecord[];
  setEsts: React.Dispatch<React.SetStateAction<EstRecord[]>>;
  targets: TargetRecord[];
  setTargets: React.Dispatch<React.SetStateAction<TargetRecord[]>>;
  month: string;
  canEdit: boolean;
  isAdmin: boolean;
}) {
  const [selectedManager, setSelectedManager] = useState<Manager>("SY");
  const [search, setSearch] = useState("");
  const normalizedSearch = search.trim().toLowerCase();
  const prevMonth = previousMonth(month);
  const prevYearMonthValue = previousYearMonth(month);

  const estMap = useMemo(() => {
    const map = new Map<string, number>();
    ests
      .filter((e) => e.month === month)
      .forEach((e) => map.set(e.storeCode, Number(e.amount || 0)));
    return map;
  }, [ests, month]);

  const targetByType = useMemo(() => {
    const totals = { store: 0, nonStore: 0 };

    targets
      .filter((t) => t.month === month && !t.storeCode)
      .forEach((t) => {
        if (t.storeType === "매장") totals.store += Number(t.amount || 0);
        else if (t.storeType === "비매장")
          totals.nonStore += Number(t.amount || 0);
      });

    return totals;
  }, [targets, month]);

  const prevEstMap = useMemo(() => {
    const map = new Map<string, number>();
    ests
      .filter((e) => e.month === prevMonth)
      .forEach((e) => map.set(e.storeCode, Number(e.amount || 0)));
    return map;
  }, [ests, prevMonth]);

  const prevSalesMap = useMemo(() => {
    const map = new Map<string, number>();
    sales
      .filter((row) => row.period === "prevMonth" && row.refMonth === month)
      .forEach((row) => map.set(row.storeCode, (map.get(row.storeCode) || 0) + Number(row.salesAmount || 0)));

    // 전월 비교용 업로드가 없고 실제 전월 데이터가 current로 저장되어 있는 경우를 대비한 보조 계산입니다.
    if (map.size === 0) {
      sales
        .filter((row) => row.period === "current" && inRange(row.saleDate, monthStart(prevMonth), monthEnd(prevMonth)))
        .forEach((row) => map.set(row.storeCode, (map.get(row.storeCode) || 0) + Number(row.salesAmount || 0)));
    }

    return map;
  }, [sales, month, prevMonth]);

  const prevYearSalesMap = useMemo(() => {
    const map = new Map<string, number>();
    sales
      .filter((row) => row.period === "prevYear" && row.refMonth === month)
      .forEach((row) =>
        map.set(
          row.storeCode,
          (map.get(row.storeCode) || 0) + Number(row.salesAmount || 0),
        ),
      );

    // 전년동월 비교용 업로드가 없고 실제 전년동월 데이터가 current로 저장되어 있는 경우를 대비한 보조 계산입니다.
    if (map.size === 0) {
      sales
        .filter(
          (row) =>
            row.period === "current" &&
            inRange(
              row.saleDate,
              monthStart(prevYearMonthValue),
              monthEnd(prevYearMonthValue),
            ),
        )
        .forEach((row) =>
          map.set(
            row.storeCode,
            (map.get(row.storeCode) || 0) + Number(row.salesAmount || 0),
          ),
        );
    }

    return map;
  }, [sales, month, prevYearMonthValue]);

  const rows = useMemo(() => {
    return stores
      .filter((store) => store.status !== "거래종료")
      .filter((store) => EST_ENTRY_MANAGERS.includes(store.manager))
      .filter((store) => store.manager === selectedManager)
      .filter((store) => {
        if (!normalizedSearch) return true;
        return [store.code, store.name, store.brand, store.channel, store.storeType]
          .some((value) => String(value || "").toLowerCase().includes(normalizedSearch));
      })
      .sort((a, b) => a.name.localeCompare(b.name, "ko-KR", { numeric: true }));
  }, [stores, selectedManager, normalizedSearch]);

  const allEntryStores = useMemo(
    () =>
      stores.filter(
        (store) =>
          store.status !== "거래종료" &&
          EST_ENTRY_MANAGERS.includes(store.manager),
      ),
    [stores],
  );

  const totalEst = rows.reduce(
    (total, store) => total + Number(estMap.get(store.code) || 0),
    0,
  );
  const storeEstTotal = allEntryStores
    .filter((store) => store.storeType === "매장")
    .reduce((total, store) => total + Number(estMap.get(store.code) || 0), 0);
  const nonStoreEstTotal = allEntryStores
    .filter((store) => store.storeType !== "매장")
    .reduce((total, store) => total + Number(estMap.get(store.code) || 0), 0);
  const managerEstTotals = EST_ENTRY_MANAGERS.map((manager) => ({
    manager,
    amount: allEntryStores
      .filter((store) => store.manager === manager)
      .reduce((total, store) => total + Number(estMap.get(store.code) || 0), 0),
  }));
  const canEditTarget = canEdit && (isAdmin || selectedManager === "SY");

  const updateEst = (store: Store, amount: number) => {
    if (!canEdit) return;
    setEsts((prev) => {
      const exists = prev.some((row) => row.month === month && row.storeCode === store.code);
      if (exists) {
        return prev.map((row) =>
          row.month === month && row.storeCode === store.code
            ? { ...row, storeName: store.name, amount }
            : row,
        );
      }
      return [...prev, { storeCode: store.code, storeName: store.name, month, amount }];
    });
  };

  const updateTargetByType = (storeType: StoreType, amount: number) => {
    if (!canEditTarget) return;
    setTargets((prev) => {
      const withoutCurrentMonthTarget = prev.filter(
        (row) =>
          row.month !== month ||
          (row.storeType !== storeType && !row.storeCode),
      );

      return [
        ...withoutCurrentMonthTarget,
        {
          storeType,
          month,
          amount,
        },
      ];
    });
  };

  return (
    <div className="space-y-4">
      <div className="rounded-2xl border border-slate-300 bg-white p-4 shadow-sm">
        <div className="flex flex-wrap items-start justify-between gap-4">
          <div className="space-y-3">
            <div className="flex flex-wrap items-center gap-2">
              {EST_ENTRY_MANAGERS.map((manager) => (
                <button
                  key={manager}
                  type="button"
                  onClick={() => setSelectedManager(manager)}
                  className={`rounded-xl px-4 py-2 text-sm font-bold ${
                    selectedManager === manager
                      ? "bg-orange-500 text-white shadow"
                      : "bg-orange-50 text-orange-900 hover:bg-orange-100"
                  }`}
                >
                  {manager}
                </button>
              ))}
            </div>
          </div>

          <div className="flex flex-wrap items-start justify-end gap-2">
            <input
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="거래처명/코드/채널 검색"
              className="h-9 w-[240px] rounded-lg border border-slate-300 bg-white px-3 text-xs outline-none focus:border-blue-500"
            />
            <div className="rounded-lg bg-orange-100 px-3 py-2 text-xs font-bold text-orange-900">
              {selectedManager} EST {won(totalEst)}
            </div>
            <div className="rounded-lg bg-emerald-100 px-3 py-2 text-xs font-bold text-emerald-900">
              <div>매장 EST {won(storeEstTotal)}</div>
              {selectedManager === "SY" && (
                <label className="mt-2 block text-[11px] font-bold text-emerald-900">
                  매장 Target
                  <input
                    type="text"
                    inputMode="numeric"
                    disabled={!canEditTarget}
                    value={targetByType.store ? won(targetByType.store) : ""}
                    onChange={(e) => updateTargetByType("매장", num(e.target.value))}
                    placeholder={canEditTarget ? "0" : "입력 기간 종료"}
                    className="mt-1 h-8 w-[150px] rounded-lg border border-emerald-200 bg-white px-2 text-right text-xs font-extrabold text-slate-900 outline-none focus:border-emerald-500 disabled:bg-slate-100 disabled:text-slate-500"
                  />
                </label>
              )}
            </div>
            <div className="rounded-lg bg-blue-100 px-3 py-2 text-xs font-bold text-blue-900">
              <div>비매장 EST {won(nonStoreEstTotal)}</div>
              {selectedManager === "SY" && (
                <label className="mt-2 block text-[11px] font-bold text-blue-900">
                  비매장 Target
                  <input
                    type="text"
                    inputMode="numeric"
                    disabled={!canEditTarget}
                    value={targetByType.nonStore ? won(targetByType.nonStore) : ""}
                    onChange={(e) => updateTargetByType("비매장", num(e.target.value))}
                    placeholder={canEditTarget ? "0" : "입력 기간 종료"}
                    className="mt-1 h-8 w-[150px] rounded-lg border border-blue-200 bg-white px-2 text-right text-xs font-extrabold text-slate-900 outline-none focus:border-blue-500 disabled:bg-slate-100 disabled:text-slate-500"
                  />
                </label>
              )}
            </div>
          </div>
        </div>
      </div>

      <div className="overflow-hidden rounded-2xl border border-slate-300 bg-white shadow-sm">
        <div className="max-h-[68vh] overflow-auto isolate">
          <table className="w-full min-w-[1450px] border-separate border-spacing-0 text-center text-[12px] whitespace-nowrap">
            <thead>
              <tr className="bg-slate-100">
                <th className="sticky top-0 z-20 border border-slate-300 bg-white px-3 py-2 font-bold text-slate-700">거래처명</th>
                <th className="sticky top-0 z-20 border border-slate-300 bg-white px-3 py-2 font-bold text-slate-700">담당자</th>
                <th className="sticky top-0 z-20 border border-slate-300 bg-white px-3 py-2 font-bold text-slate-700">채널</th>
                <th className="sticky top-0 z-20 border border-slate-300 bg-[#F7FCEB] px-3 py-2 font-bold text-black">전년동월 매출</th>
                <th className="sticky top-0 z-20 border border-slate-300 bg-[#F3FAFD] px-3 py-2 font-bold text-black">전월 매출</th>
                <th className="sticky top-0 z-20 border border-slate-300 bg-[#1E3A5F] px-3 py-2 text-[14px] font-bold text-white">전월 EST</th>
                <th className="sticky top-0 z-20 border border-slate-300 bg-[#1E3A5F] px-3 py-2 text-[14px] font-bold text-white">전월 EST 달성률</th>
                <th className="sticky top-0 z-20 border border-slate-300 bg-yellow-100 px-3 py-2 text-[14px] font-bold text-black">{month} EST 입력</th>
                <th className="sticky top-0 z-20 border border-slate-300 bg-yellow-100 px-3 py-2 text-[14px] font-bold text-black">전월 EST 대비 차이</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((store) => {
                const value = estMap.get(store.code) || 0;
                const prevYearSales = prevYearSalesMap.get(store.code) || 0;
                const prevSales = prevSalesMap.get(store.code) || 0;
                const prevEst = prevEstMap.get(store.code) || 0;
                const prevEstRate = prevEst ? (prevSales / prevEst) * 100 : 0;
                const estDifference = Number(value || 0) - Number(prevEst || 0);
                const isCriticalRate = Boolean(prevEst && prevEstRate < 30);
                const rateTone = !prevEst
                  ? "bg-slate-100 text-slate-500 ring-slate-200"
                  : isCriticalRate
                    ? "bg-red-600 text-white ring-red-300 shadow-red-300 animate-pulse"
                    : prevEstRate > 120
                      ? "bg-yellow-100 text-yellow-900 ring-yellow-300"
                      : prevEstRate >= 80
                        ? "bg-emerald-100 text-emerald-800 ring-emerald-300"
                        : "bg-slate-100 text-slate-700 ring-slate-200";
                const differenceTone = estDifference > 0
                  ? "text-emerald-700 bg-emerald-50"
                  : estDifference < 0
                    ? "text-red-600 bg-red-50"
                    : "text-slate-500 bg-slate-50";
                return (
                  <tr key={store.code} className="hover:bg-orange-50/60">
                    <td className="border border-slate-300 px-3 py-2 text-center font-semibold text-slate-900">{store.name}</td>
                    <td className="border border-slate-300 px-3 py-2 font-bold text-slate-900">{store.manager}</td>
                    <td className="border border-slate-300 px-3 py-2 text-slate-700">{store.storeType === "매장" ? "매장" : "비매장"}</td>
                    <td className="border border-slate-300 px-3 py-2 text-right font-semibold text-slate-900">{won(prevYearSales)}</td>
                    <td className="border border-slate-300 px-3 py-2 text-right font-semibold text-slate-900">{won(prevSales)}</td>
                    <td className="border border-slate-300 px-3 py-2 text-right font-semibold text-slate-900">{won(prevEst)}</td>
                    <td className="border border-slate-300 px-3 py-2">
                      <span
                        className={`inline-flex min-w-[86px] justify-center rounded-full px-3 py-1.5 text-xs font-extrabold ring-2 shadow-sm ${rateTone}`}
                        style={isCriticalRate ? { animationDuration: "0.65s" } : undefined}
                      >
                        {isCriticalRate ? "🚨 " : ""}{prevEst ? pct(prevEstRate) : "-"}
                      </span>
                    </td>
                    <td className="border border-slate-300 px-3 py-2">
                      <input
                        type="text"
                        inputMode="numeric"
                        disabled={!canEdit}
                        value={value ? won(value) : ""}
                        onChange={(e) => updateEst(store, num(e.target.value))}
                        placeholder="0"
                        className="h-9 w-full min-w-[150px] rounded-lg border border-slate-300 bg-white px-3 text-right text-sm font-bold text-slate-900 outline-none focus:border-orange-500 disabled:bg-slate-100 disabled:text-slate-500"
                      />
                    </td>
                    <td className={`border border-slate-300 px-3 py-2 text-right font-extrabold ${differenceTone}`}>
                      {value || prevEst
                        ? `${estDifference > 0 ? "+" : ""}${won(estDifference)}`
                        : "-"}
                    </td>
                  </tr>
                );
              })}
              {!rows.length && (
                <tr>
                  <td colSpan={9} className="border border-slate-300 p-8 text-center text-slate-500">
                    표시할 거래처가 없습니다.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function ItemCostStatus({
  sales,
  itemCosts,
  setItemCosts,
  isAdmin,
}: {
  sales: SalesRecord[];
  itemCosts: ItemCostRecord[];
  setItemCosts: React.Dispatch<React.SetStateAction<ItemCostRecord[]>>;
  isAdmin: boolean;
}) {
  const [search, setSearch] = useState("");
  const [showChangedOnly, setShowChangedOnly] = useState(false);
  const [openItemCode, setOpenItemCode] = useState("");
  const normalizedSearch = search.trim().toLowerCase();

  useEffect(() => {
    const todayText = today();
    setItemCosts((prev) => {
      let changed = false;
      const next = prev.map((item) => {
        const scheduledCost = Number(item.nextCost || 0);
        const shouldApply =
          item.effectiveDate &&
          item.effectiveDate <= todayText &&
          scheduledCost > 0 &&
          scheduledCost !== Number(item.currentCost || 0);
        if (!shouldApply) return item;

        const historyId = `${item.itemCode}|${item.effectiveDate}|${scheduledCost}`;
        const alreadyRecorded = (item.history || []).some(
          (h) => h.id === historyId,
        );
        changed = true;
        return {
          ...item,
          currentCost: scheduledCost,
          nextCost: undefined,
          effectiveDate: undefined,
          history: alreadyRecorded
            ? item.history || []
            : [
                {
                  id: historyId,
                  changedAt: todayText,
                  effectiveDate: item.effectiveDate || todayText,
                  previousCost: Number(item.currentCost || 0),
                  newCost: scheduledCost,
                  memo: item.memo || "예정 매입가 자동 적용",
                },
                ...(item.history || []),
              ],
        };
      });
      return changed ? next : prev;
    });
  }, [setItemCosts]);

  const itemBaseRows = useMemo(() => {
    const map = new Map<
      string,
      { itemCode: string; itemName: string; estimatedCost: number }
    >();
    sales.forEach((row) => {
      const key = row.itemCode || row.itemName || "미지정";
      if (!map.has(key)) {
        const quantity = Number(row.quantity || 0);
        map.set(key, {
          itemCode: row.itemCode || "-",
          itemName: row.itemName || "미지정",
          estimatedCost: quantity ? Number(row.costAmount || 0) / quantity : 0,
        });
      }
    });
    itemCosts.forEach((item) => {
      const key = item.itemCode || item.itemName || "미지정";
      if (!map.has(key)) {
        map.set(key, {
          itemCode: item.itemCode || "-",
          itemName: item.itemName || "미지정",
          estimatedCost: Number(item.currentCost || 0),
        });
      }
    });
    return Array.from(map.values()).sort((a, b) =>
      a.itemName.localeCompare(b.itemName, "ko-KR", { numeric: true }),
    );
  }, [sales, itemCosts]);

  const costMap = useMemo(
    () => new Map(itemCosts.map((item) => [item.itemCode, item])),
    [itemCosts],
  );

  const rows = useMemo(() => {
    return itemBaseRows
      .map((base) => {
        const saved = costMap.get(base.itemCode);
        const currentCost = Number(
          saved?.currentCost || base.estimatedCost || 0,
        );
        const nextCost = Number(saved?.nextCost || 0);
        const diff = nextCost ? nextCost - currentCost : 0;
        const rate = currentCost ? (diff / currentCost) * 100 : 0;
        const dday = saved?.effectiveDate
          ? daysBetween(today(), saved.effectiveDate)
          : null;
        const status = !saved?.effectiveDate
          ? "-"
          : dday !== null && dday > 0
            ? "예정"
            : "적용완료";
        return {
          ...base,
          saved,
          currentCost,
          nextCost,
          effectiveDate: saved?.effectiveDate || "",
          memo: saved?.memo || "",
          history: saved?.history || [],
          diff,
          rate,
          dday,
          status,
        };
      })
      .filter((row) => {
        if (showChangedOnly && !row.nextCost && !row.history.length) return false;
        if (!normalizedSearch) return true;
        return [row.itemCode, row.itemName, row.memo]
          .some((value) => String(value || "").toLowerCase().includes(normalizedSearch));
      });
  }, [itemBaseRows, costMap, normalizedSearch, showChangedOnly]);

  const upsertItemCost = (
    base: { itemCode: string; itemName: string; estimatedCost: number },
    patch: Partial<ItemCostRecord>,
  ) => {
    if (!isAdmin) return;
    setItemCosts((prev) => {
      const exists = prev.some((item) => item.itemCode === base.itemCode);
      if (exists) {
        return prev.map((item) =>
          item.itemCode === base.itemCode
            ? {
                ...item,
                itemName: base.itemName,
                history: item.history || [],
                ...patch,
              }
            : item,
        );
      }
      return [
        ...prev,
        {
          itemCode: base.itemCode,
          itemName: base.itemName,
          currentCost: Number(base.estimatedCost || 0),
          history: [],
          ...patch,
        },
      ];
    });
  };

  const applyNow = (row: (typeof rows)[number]) => {
    if (!isAdmin || !row.nextCost) return;
    const todayText = today();
    const historyId = `${row.itemCode}|${row.effectiveDate || todayText}|${row.nextCost}`;
    upsertItemCost(row, {
      currentCost: row.nextCost,
      nextCost: undefined,
      effectiveDate: undefined,
      history: [
        {
          id: historyId,
          changedAt: todayText,
          effectiveDate: row.effectiveDate || todayText,
          previousCost: row.currentCost,
          newCost: row.nextCost,
          memo: row.memo || "매입가 변경",
        },
        ...row.history.filter((h) => h.id !== historyId),
      ],
    });
  };

  const deleteHistory = (row: (typeof rows)[number], historyId: string) => {
    if (!isAdmin) return;
    if (!window.confirm("선택한 매입가 변경 이력을 삭제할까요?")) return;
    upsertItemCost(row, {
      history: row.history.filter((h) => h.id !== historyId),
    });
  };

  return (
    <div className="flex min-h-0 flex-1 flex-col space-y-3">
      <div className="bg-white py-1">
        <div className="flex flex-wrap items-end justify-end gap-3">
          <div>
            <div className="text-base font-extrabold text-slate-900">매입가 정보</div>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <input
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="품목코드/품목명/메모 검색"
              className="h-9 w-[260px] rounded-lg border border-slate-300 bg-white px-3 text-xs outline-none focus:border-blue-500"
            />
            <label className="flex h-9 items-center gap-2 rounded-lg border border-slate-300 bg-slate-50 px-3 text-xs font-bold text-slate-700">
              <input
                type="checkbox"
                checked={showChangedOnly}
                onChange={(e) => setShowChangedOnly(e.target.checked)}
              />
              변동 품목만
            </label>
          </div>
        </div>
      </div>

      <div className="overflow-hidden rounded-2xl border border-slate-300 bg-white shadow-sm">
        <div className="max-h-[78vh] overflow-auto isolate">
          <table className="w-full min-w-[1350px] border-separate border-spacing-0 text-center text-[12px] whitespace-nowrap">
            <thead>
              <tr className="bg-slate-100">
                <th className="sticky top-0 z-20 border border-slate-300 bg-slate-100 px-3 py-2 font-bold text-slate-700">품목코드</th>
                <th className="sticky top-0 z-20 border border-slate-300 bg-slate-100 px-3 py-2 font-bold text-slate-700">품목명</th>
                <th className="sticky top-0 z-20 border border-slate-300 bg-[#F3FAFD] px-3 py-2 font-bold text-black">현재 매입가</th>
                <th className="sticky top-0 z-20 border border-slate-300 bg-orange-50 px-3 py-2 font-bold text-orange-800">다음 매입가</th>
                <th className="sticky top-0 z-20 border border-slate-300 bg-orange-50 px-3 py-2 font-bold text-orange-800">적용 예정일</th>
                <th className="sticky top-0 z-20 border border-slate-300 bg-slate-100 px-3 py-2 font-bold text-slate-700">상태</th>
                <th className="sticky top-0 z-20 border border-slate-300 bg-slate-100 px-3 py-2 font-bold text-slate-700">D-Day</th>
                <th className="sticky top-0 z-20 border border-slate-300 bg-slate-100 px-3 py-2 font-bold text-slate-700">변동금액</th>
                <th className="sticky top-0 z-20 border border-slate-300 bg-slate-100 px-3 py-2 font-bold text-slate-700">변동률</th>
                <th className="sticky top-0 z-20 border border-slate-300 bg-slate-100 px-3 py-2 font-bold text-slate-700">메모</th>
                <th className="sticky top-0 z-20 border border-slate-300 bg-slate-100 px-3 py-2 font-bold text-slate-700">History</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((row) => {
                const isDueSoon = row.dday !== null && row.dday <= 7 && row.dday >= 0;
                const isUpcoming = row.dday !== null && row.dday > 7 && row.dday <= 30;
                const diffTone = row.diff > 0 ? "text-red-600" : row.diff < 0 ? "text-blue-600" : "text-slate-500";
                return (
                  <Fragment key={row.itemCode}>
                    <tr className={isDueSoon ? "bg-red-50 hover:bg-red-100" : isUpcoming ? "bg-yellow-50 hover:bg-yellow-100" : "hover:bg-slate-50"}>
                      <td className="border border-slate-300 px-3 py-2 text-slate-700">{row.itemCode}</td>
                      <td className="border border-slate-300 px-3 py-2 text-left font-semibold text-slate-900">{row.itemName}</td>
                      <td className="border border-slate-300 px-3 py-2">
                        <input
                          type="text"
                          inputMode="numeric"
                          disabled={!isAdmin}
                          value={row.currentCost ? won(row.currentCost) : ""}
                          onChange={(e) => upsertItemCost(row, { currentCost: num(e.target.value) })}
                          placeholder="0"
                          className="h-9 w-full min-w-[120px] rounded-lg border border-slate-300 bg-white px-3 text-right text-sm font-bold outline-none focus:border-blue-500 disabled:bg-slate-100 disabled:text-slate-500"
                        />
                      </td>
                      <td className="border border-slate-300 px-3 py-2">
                        <input
                          type="text"
                          inputMode="numeric"
                          disabled={!isAdmin}
                          value={row.nextCost ? won(row.nextCost) : ""}
                          onChange={(e) => upsertItemCost(row, { nextCost: num(e.target.value) || undefined })}
                          placeholder="변경 예정가"
                          className="h-9 w-full min-w-[120px] rounded-lg border border-orange-300 bg-orange-50 px-3 text-right text-sm font-bold outline-none focus:border-orange-500 disabled:bg-slate-100 disabled:text-slate-500"
                        />
                      </td>
                      <td className="border border-slate-300 px-3 py-2">
                        <input
                          type="date"
                          disabled={!isAdmin}
                          value={row.effectiveDate}
                          onChange={(e) => upsertItemCost(row, { effectiveDate: e.target.value || undefined })}
                          className="h-9 rounded-lg border border-slate-300 bg-white px-2 text-xs font-bold outline-none focus:border-orange-500 disabled:bg-slate-100 disabled:text-slate-500"
                        />
                      </td>
                      <td className="border border-slate-300 px-3 py-2">
                        <span className={`rounded-full px-3 py-1 text-xs font-extrabold ${row.status === "예정" ? "bg-yellow-100 text-yellow-800" : row.status === "적용완료" ? "bg-emerald-100 text-emerald-800" : "bg-slate-100 text-slate-500"}`}>
                          {row.status}
                        </span>
                      </td>
                      <td className="border border-slate-300 px-3 py-2 font-bold">
                        {row.dday === null ? "-" : row.dday > 0 ? `D-${row.dday}` : row.dday === 0 ? "D-Day" : "적용완료"}
                      </td>
                      <td className={`border border-slate-300 px-3 py-2 text-right font-extrabold ${diffTone}`}>
                        {row.diff > 0 ? "▲ " : row.diff < 0 ? "▼ " : ""}{row.diff ? won(Math.abs(row.diff)) : "-"}
                      </td>
                      <td className={`border border-slate-300 px-3 py-2 text-right font-bold ${diffTone}`}>
                        {row.nextCost ? pct(row.rate) : "-"}
                      </td>
                      <td className="border border-slate-300 px-3 py-2">
                        <input
                          type="text"
                          disabled={!isAdmin}
                          value={row.memo}
                          onChange={(e) => upsertItemCost(row, { memo: e.target.value })}
                          placeholder="예: 공급가 인상, 환율 반영"
                          className="h-9 w-full min-w-[180px] rounded-lg border border-slate-300 bg-white px-3 text-xs outline-none focus:border-blue-500 disabled:bg-slate-100 disabled:text-slate-500"
                        />
                      </td>
                      <td className="border border-slate-300 px-3 py-2">
                        <div className="flex items-center justify-center gap-2">
                          {isAdmin && row.nextCost > 0 && (
                            <button
                              type="button"
                              onClick={() => applyNow(row)}
                              className="rounded-lg bg-orange-600 px-2 py-1 text-[11px] font-bold text-white hover:bg-orange-700"
                            >
                              적용
                            </button>
                          )}
                          <button
                            type="button"
                            onClick={() => setOpenItemCode(openItemCode === row.itemCode ? "" : row.itemCode)}
                            className="rounded-lg bg-slate-800 px-2 py-1 text-[11px] font-bold text-white hover:bg-slate-700"
                          >
                            {openItemCode === row.itemCode ? "닫기" : `보기 ${row.history.length}`}
                          </button>
                        </div>
                      </td>
                    </tr>
                    {openItemCode === row.itemCode && (
                      <tr>
                        <td colSpan={11} className="border border-slate-300 bg-slate-50 p-3">
                          <div className="text-left text-xs font-bold text-slate-700">{row.itemName} 매입가 History</div>
                          <div className="mt-2 overflow-auto rounded-xl border border-slate-300 bg-white">
                            <table className="w-full min-w-[760px] text-center text-[11px]">
                              <thead className="bg-slate-100">
                                <tr>
                                  <th className="border border-slate-300 px-2 py-1">변경입력일</th>
                                  <th className="border border-slate-300 px-2 py-1">적용일</th>
                                  <th className="border border-slate-300 px-2 py-1">이전가</th>
                                  <th className="border border-slate-300 px-2 py-1">변경가</th>
                                  <th className="border border-slate-300 px-2 py-1">증감</th>
                                  <th className="border border-slate-300 px-2 py-1">메모</th>
                                  {isAdmin && <th className="border border-slate-300 px-2 py-1">관리</th>}
                                </tr>
                              </thead>
                              <tbody>
                                {row.history.map((h) => (
                                  <tr key={h.id}>
                                    <td className="border border-slate-300 px-2 py-1">{h.changedAt}</td>
                                    <td className="border border-slate-300 px-2 py-1">{h.effectiveDate}</td>
                                    <td className="border border-slate-300 px-2 py-1 text-right">{won(h.previousCost)}</td>
                                    <td className="border border-slate-300 px-2 py-1 text-right font-bold">{won(h.newCost)}</td>
                                    <td className={`border border-slate-300 px-2 py-1 text-right font-bold ${h.newCost - h.previousCost >= 0 ? "text-red-600" : "text-blue-600"}`}>
                                      {h.newCost - h.previousCost >= 0 ? "+" : ""}{won(h.newCost - h.previousCost)}
                                    </td>
                                    <td className="border border-slate-300 px-2 py-1 text-left">{h.memo || "-"}</td>
                                    {isAdmin && (
                                      <td className="border border-slate-300 px-2 py-1">
                                        <button
                                          type="button"
                                          onClick={() => deleteHistory(row, h.id)}
                                          className="rounded-md bg-red-50 px-2 py-1 text-[11px] font-bold text-red-600 hover:bg-red-100"
                                        >
                                          삭제
                                        </button>
                                      </td>
                                    )}
                                  </tr>
                                ))}
                                {!row.history.length && (
                                  <tr>
                                    <td colSpan={isAdmin ? 7 : 6} className="border border-slate-300 p-5 text-center text-slate-500">
                                      아직 저장된 변경 이력이 없습니다.
                                    </td>
                                  </tr>
                                )}
                              </tbody>
                            </table>
                          </div>
                        </td>
                      </tr>
                    )}
                  </Fragment>
                );
              })}
              {!rows.length && (
                <tr>
                  <td colSpan={11} className="border border-slate-300 p-8 text-center text-slate-500">
                    표시할 품목이 없습니다.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function HeaderTimeInfo({ title, value }: { title: string; value: number }) {
  return (
    <div className="min-w-[76px] rounded-lg border border-gray-300/70 bg-white/70 px-3 py-1.5 text-center shadow-sm">
      <p className="text-[11px] font-semibold text-slate-500">{title}</p>
      <p className="mt-0.5 text-[16px] font-bold text-slate-900">
        {Number.isInteger(value) ? value : value.toFixed(1)}일
      </p>
    </div>
  );
}

function HeaderMetricInfo({
  title,
  value,
}: {
  title: string;
  value: string | number;
}) {
  return (
    <div className="min-w-[76px] rounded-lg border border-gray-300/70 bg-white/70 px-3 py-1.5 text-center shadow-sm">
      <p className="text-[11px] font-semibold text-slate-500">{title}</p>
      <p className="mt-0.5 text-[16px] font-bold text-slate-900">{value}</p>
    </div>
  );
}

function KpiGroup({
  items,
  className = "bg-slate-50/75",
}: {
  items: {
    title: string;
    value: string | number;
    color?: string;
    format?: "won" | "percent" | "number";
    highlightClass?: string;
  }[];
  className?: string;
}) {
  return (
    <div className={`h-full rounded-xl border border-gray-300/70 p-2.5 shadow-sm backdrop-blur ${className}`}>
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
            <div
              key={item.title}
              className={`flex min-h-[31px] flex-wrap items-center justify-between gap-x-1.5 gap-y-0.5 rounded-lg px-1.5 py-0.5 first:pt-0 last:pb-0 ${item.highlightClass || ""}`}
            >
              <p className="shrink-0 break-keep text-[13px] font-semibold text-black">
                {item.title}
              </p>
              <p
                className={`min-w-0 max-w-full flex-1 whitespace-normal break-all text-right text-[16px] font-bold leading-tight tracking-tight md:text-[16px] xl:text-[17px] 2xl:text-[18px] ${item.color || "text-black"}`}
              >
                {value}
              </p>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function metricsByStoreType(
  stores: Store[],
  targets: TargetRecord[],
  ests: EstRecord[],
  month: string,
) {
  const storeCodes = new Set(
    stores.filter((s) => s.storeType === "매장").map((s) => s.code),
  );

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
  const storeEst = ests
    .filter((e) => e.month === month && storeCodes.has(e.storeCode))
    .reduce((a, b) => a + b.amount, 0);
  const nonStoreEst = ests
    .filter((e) => e.month === month && !storeCodes.has(e.storeCode))
    .reduce((a, b) => a + b.amount, 0);
  return { storeTarget, nonStoreTarget, storeEst, nonStoreEst };
}

function SalesTargetHeaderKpi({
  stores,
  sales,
  targets,
  ests,
  month,
  date,
}: {
  stores: Store[];
  sales: SalesRecord[];
  targets: TargetRecord[];
  ests: EstRecord[];
  month: string;
  date: string;
}) {
  const { storeTarget, nonStoreTarget } = metricsByStoreType(
    stores,
    targets,
    ests,
    month,
  );
  const targetTotal = storeTarget + nonStoreTarget;
  const currentSales = sales
    .filter(
      (row) =>
        row.period === "current" &&
        inRange(row.saleDate, monthStart(month), date),
    )
    .reduce((total, row) => total + Number(row.salesAmount || 0), 0);
  const rate = targetTotal ? (currentSales / targetTotal) * 100 : 0;

  return (
    <div className="ml-auto min-w-[430px] rounded-xl border border-yellow-300 bg-yellow-100 px-4 py-2 shadow-sm">
      <div className="grid grid-cols-[1.25fr_1fr_1fr_1fr] items-center gap-2 text-black">
        <div className="rounded-lg bg-yellow-200 px-3 py-2 text-center">
          <div className="text-[12px] font-extrabold">TARGET 달성률</div>
          <div className="mt-0.5 text-[30px] font-black leading-none">{targetTotal ? pct(rate) : "-"}</div>
        </div>
        <div className="text-center">
          <div className="text-[11px] font-bold">매장 Target</div>
          <div className="mt-1 text-[16px] font-black">{won(storeTarget)}</div>
        </div>
        <div className="text-center">
          <div className="text-[11px] font-bold">비매장 Target</div>
          <div className="mt-1 text-[16px] font-black">{won(nonStoreTarget)}</div>
        </div>
        <div className="text-center">
          <div className="text-[11px] font-bold">총 Target</div>
          <div className="mt-1 text-[16px] font-black">{won(targetTotal)}</div>
        </div>
      </div>
    </div>
  );
}
function DashboardTopKpis({
  stores,
  sales,
  targets,
  ests,
  itemCosts,
  month,
  date,
}: {
  stores: Store[];
  sales: SalesRecord[];
  targets: TargetRecord[];
  ests: EstRecord[];
  itemCosts: ItemCostRecord[];
  month: string;
  date: string;
}) {
  const [costAlertOpen, setCostAlertOpen] = useState(false);
  const current = sales.filter(
    (s) =>
      s.period === "current" && inRange(s.saleDate, monthStart(month), date),
  );
  const currentFullMonth = sales.filter(
    (s) =>
      s.period === "current" &&
      inRange(s.saleDate, monthStart(month), monthEnd(month)),
  );
  const stMap = storeMap(stores);
  const storeTypeOf = (row: SalesRecord) =>
    stMap.get(row.storeCode)?.storeType || row.storeType || "비매장";
  const currentStoreSales = current
    .filter((s) => storeTypeOf(s) === "매장")
    .reduce((total, row) => total + Number(row.salesAmount || 0), 0);
  const currentNonStoreSales = current
    .filter((s) => storeTypeOf(s) !== "매장")
    .reduce((total, row) => total + Number(row.salesAmount || 0), 0);
  const nextMonth = addMonths(month, 1);
  const nextMonthNonStoreOrderAmount = sales
    .filter(
      (s) =>
        s.period === "current" &&
        inRange(s.saleDate, monthStart(nextMonth), monthEnd(nextMonth)) &&
        storeTypeOf(s) !== "매장",
    )
    .reduce((total, row) => total + Number(row.salesAmount || 0), 0);
  const currentSales = sum(current, "salesAmount");
  const fullMonthSales = sum(currentFullMonth, "salesAmount");
  const profitAmount = sum(current, "profitAmount");
  const profitRate = weightedProfitRate(current);
  const { storeTarget, nonStoreTarget, storeEst, nonStoreEst } =
    metricsByStoreType(stores, targets, ests, month);
  const targetTotal = storeTarget + nonStoreTarget;
  const estTotal = storeEst + nonStoreEst;

  const scheduledCostItems = itemCosts.filter(
    (item) => item.effectiveDate && Number(item.nextCost || 0) > 0,
  );
  const overdueCount = scheduledCostItems.filter((item) => {
    const dday = daysBetween(today(), item.effectiveDate || today());
    return dday < 0;
  }).length;
  const dueTodayCount = scheduledCostItems.filter((item) => {
    const dday = daysBetween(today(), item.effectiveDate || today());
    return dday === 0;
  }).length;
  const dueSoonCount = scheduledCostItems.filter((item) => {
    const dday = daysBetween(today(), item.effectiveDate || today());
    return dday > 0 && dday <= 7;
  }).length;
  const upcomingCount = scheduledCostItems.filter((item) => {
    const dday = daysBetween(today(), item.effectiveDate || today());
    return dday > 7 && dday <= 30;
  }).length;
  const recentHistoryCount = itemCosts.reduce((total, item) => {
    return (
      total +
      (item.history || []).filter((history) => {
        const diff = daysBetween(history.changedAt, today());
        return diff >= 0 && diff <= 30;
      }).length
    );
  }, 0);

  const currentItemSummary = new Map<string, { itemCode: string; itemName: string; sales: number; cost: number; quantity: number }>();
  current.forEach((row) => {
    const key = row.itemCode || row.itemName || "미지정";
    const prev = currentItemSummary.get(key) || { itemCode: row.itemCode || "-", itemName: row.itemName || "미지정", sales: 0, cost: 0, quantity: 0 };
    prev.sales += Number(row.salesAmount || 0);
    prev.cost += Number(row.costAmount || 0);
    prev.quantity += Number(row.quantity || 0);
    currentItemSummary.set(key, prev);
  });
  const costAlertRows = scheduledCostItems
    .map((item) => {
      const summary = currentItemSummary.get(item.itemCode) || { itemCode: item.itemCode, itemName: item.itemName, sales: 0, cost: 0, quantity: 0 };
      const previousUnitCost = summary.quantity ? summary.cost / summary.quantity : Number(item.currentCost || 0);
      const sellingUnitPrice = summary.quantity ? summary.sales / summary.quantity : 0;
      const nextCost = Number(item.nextCost || 0);
      const currentMarginRate = sellingUnitPrice ? ((sellingUnitPrice - previousUnitCost) / sellingUnitPrice) * 100 : 0;
      const nextMarginRate = sellingUnitPrice ? ((sellingUnitPrice - nextCost) / sellingUnitPrice) * 100 : 0;
      return { ...item, previousUnitCost, sellingUnitPrice, currentMarginRate, nextMarginRate, nextCost };
    })
    .sort((a, b) => String(a.effectiveDate || "").localeCompare(String(b.effectiveDate || "")));

  const activeCostAlertCount = overdueCount + dueTodayCount + dueSoonCount + upcomingCount;

  return (
    <>
    <div className="grid w-full grid-cols-1 gap-2 lg:grid-cols-5">
      <KpiGroup
        items={[
          { title: "매장 Target", value: storeTarget, format: "won" },
          { title: "비매장 Target", value: nonStoreTarget, format: "won" },
          {
            title: "총 Target",
            value: targetTotal,
            format: "won",
            color: "text-slate-900",
          },
          {
            title: "Target 달성률",
            value: targetTotal ? (currentSales / targetTotal) * 100 : 0,
            format: "percent",
            color: "text-black",
            highlightClass: "bg-[#FFF9D9] ring-1 ring-[#F3E08A]",
          },
        ]}
      />
      <KpiGroup
        className="est-summary-card"
        items={[
          { title: "매장 EST", value: storeEst, format: "won" },
          { title: "비매장 EST", value: nonStoreEst, format: "won" },
          {
            title: "총 EST",
            value: estTotal,
            format: "won",
            color: "text-slate-900",
          },
          {
            title: "EST 달성률",
            value: estTotal ? (currentSales / estTotal) * 100 : 0,
            format: "percent",
            color: "text-slate-900",
          },
        ]}
      />
      <KpiGroup
        items={[
          {
            title: "당월 매장 매출",
            value: currentStoreSales,
            format: "won",
            color: "text-slate-900",
          },
          {
            title: "당월 비매장 매출",
            value: currentNonStoreSales,
            format: "won",
            color: "text-slate-900",
          },
          { title: "당일까지 매출", value: currentSales, format: "won" },
          { title: "당월 전체 매출", value: fullMonthSales, format: "won" },
        ]}
      />
      <KpiGroup
        items={[
          {
            title: "이익금액",
            value: profitAmount,
            format: "won",
            color: "text-slate-900",
          },
          {
            title: "이익률",
            value: profitRate,
            format: "percent",
            color: "text-slate-900",
          },
          {
            title: "익월 비매장 발주 접수액",
            value: nextMonthNonStoreOrderAmount,
            format: "won",
            color: "text-slate-900",
          },
        ]}
      />
      <button
        type="button"
        onClick={() => activeCostAlertCount > 0 && setCostAlertOpen(true)}
        className={`h-full rounded-xl border border-gray-200 bg-white p-2.5 text-left shadow-sm ${activeCostAlertCount > 0 ? "cursor-pointer hover:bg-orange-50 hover:shadow-md" : "cursor-default"}`}
      >
        <div className="divide-y divide-gray-200 text-black">
          <div className="flex min-h-[31px] items-center justify-between px-1.5"><span className="text-[13px] font-semibold">매입가 알림</span><span className="text-[18px] font-black">{activeCostAlertCount ? `${activeCostAlertCount}건` : "정상"}</span></div>
          <div className="flex min-h-[31px] items-center justify-between px-1.5"><span className="text-[13px] font-semibold">오늘 적용</span><span className="text-[16px] font-bold">{dueTodayCount}건{overdueCount ? ` / 지남 ${overdueCount}건` : ""}</span></div>
          <div className="flex min-h-[31px] items-center justify-between px-1.5"><span className="text-[13px] font-semibold">7일 이내</span><span className="text-[16px] font-bold">{dueSoonCount}건</span></div>
          <div className="flex min-h-[31px] items-center justify-between px-1.5"><span className="text-[13px] font-semibold">30일 이내</span><span className="text-[16px] font-bold">{upcomingCount}건 · 최근변경 {recentHistoryCount}건</span></div>
        </div>
      </button>
    </div>
    {costAlertOpen && (
      <div className="fixed inset-0 z-[100] flex items-center justify-center bg-black/45 p-4" onMouseDown={() => setCostAlertOpen(false)}>
        <div className="max-h-[82vh] w-full max-w-6xl overflow-hidden rounded-2xl bg-white shadow-2xl" onMouseDown={(e) => e.stopPropagation()}>
          <div className="flex items-center justify-between border-b border-gray-200 px-5 py-4">
            <div><h3 className="text-lg font-extrabold text-black">매입가 인상 상세 품목</h3><p className="mt-1 text-xs text-black">당월 매출의 판매가·원가를 기준으로 계산합니다.</p></div>
            <button type="button" onClick={() => setCostAlertOpen(false)} className="rounded-lg border border-gray-300 bg-white px-3 py-1.5 text-xs font-bold text-black">닫기</button>
          </div>
          <div className="max-h-[68vh] overflow-auto bg-white">
            <table className="w-full min-w-[1050px] border-separate border-spacing-0 text-center text-[12px] text-black">
              <thead><tr>
                <th className="sticky top-0 z-20 border border-gray-200 bg-white px-3 py-2">품목코드</th>
                <th className="sticky top-0 z-20 border border-gray-200 bg-white px-3 py-2">품목명</th>
                <th className="sticky top-0 z-20 border border-gray-200 bg-blue-100 px-3 py-2">전 매입가</th>
                <th className="sticky top-0 z-20 border border-gray-200 bg-blue-100 px-3 py-2">현재 판매가</th>
                <th className="sticky top-0 z-20 border border-gray-200 bg-orange-100 px-3 py-2">현재 판매가 대비 이익률</th>
                <th className="sticky top-0 z-20 border border-gray-200 bg-yellow-100 px-3 py-2">인상 후 매입가</th>
                <th className="sticky top-0 z-20 border border-gray-200 bg-orange-100 px-3 py-2">인상 후 판매가 이익률</th>
                <th className="sticky top-0 z-20 border border-gray-200 bg-white px-3 py-2">적용 예정일</th>
              </tr></thead>
              <tbody>{costAlertRows.map((row) => (
                <tr key={row.itemCode}>
                  <td className="border border-gray-200 px-3 py-2">{row.itemCode}</td>
                  <td className="border border-gray-200 px-3 py-2 text-left font-semibold">{row.itemName}</td>
                  <td className="border border-gray-200 px-3 py-2 text-right font-bold">{won(row.previousUnitCost)}</td>
                  <td className="border border-gray-200 px-3 py-2 text-right font-bold">{won(row.sellingUnitPrice)}</td>
                  <td className="border border-gray-200 px-3 py-2 text-right font-bold">{pct(row.currentMarginRate)}</td>
                  <td className="border border-gray-200 px-3 py-2 text-right font-bold">{won(row.nextCost)}</td>
                  <td className="border border-gray-200 px-3 py-2 text-right font-bold">{pct(row.nextMarginRate)}</td>
                  <td className="border border-gray-200 px-3 py-2">{row.effectiveDate || "-"}</td>
                </tr>
              ))}{!costAlertRows.length && <tr><td colSpan={8} className="border border-gray-200 p-8">표시할 매입가 알림이 없습니다.</td></tr>}</tbody>
            </table>
          </div>
        </div>
      </div>
    )}
    </>
  );
}

function Dashboard({
  stores,
  sales,
  targets,
  ests,
  month,
  date,
  timeGone,
  codeMappings,
}: {
  stores: Store[];
  sales: SalesRecord[];
  targets: TargetRecord[];
  ests: EstRecord[];
  month: string;
  date: string;
  timeGone: ReturnType<typeof getTimeGone>;
  codeMappings: StoreCodeMapping[];
}) {
  void targets;
  void timeGone;
  void codeMappings;

  const stMap = storeMap(stores);
  const normalizeDashboardManager = (value: unknown) =>
    norm(value).trim().toUpperCase();
  const managerOfSale = (row: SalesRecord) => {
    const store = stMap.get(row.storeCode);
    return store
      ? normalizeDashboardManager(store.manager) || "미지정"
      : "미지정";
  };
  const managerOfEst = (row: EstRecord) => {
    const store = stMap.get(row.storeCode);
    return store
      ? normalizeDashboardManager(store.manager) || "미지정"
      : "미지정";
  };

  const current = sales.filter(
    (s) =>
      s.period === "current" && inRange(s.saleDate, monthStart(month), date),
  );
  const currentFullMonth = sales.filter(
    (s) =>
      s.period === "current" &&
      inRange(s.saleDate, monthStart(month), monthEnd(month)),
  );
  const prevMonth = sales.filter(
    (s) => s.period === "prevMonth" && s.refMonth === month,
  );
  const prevYear = sales.filter(
    (s) => s.period === "prevYear" && s.refMonth === month,
  );

  const managers = Array.from(
    new Set(
      stores
        .map((store) => normalizeDashboardManager(store.manager))
        .filter(Boolean),
    ),
  ).sort((a, b) => a.localeCompare(b, "ko"));
  const hasUnassigned =
    [...currentFullMonth, ...prevMonth, ...prevYear].some(
      (row) => managerOfSale(row) === "미지정",
    ) ||
    ests.some(
      (row) => row.month === month && managerOfEst(row) === "미지정",
    );
  if (hasUnassigned) managers.push("미지정");

  const rows = managers.map((manager) => {
    const currentRows = current.filter((s) => managerOfSale(s) === manager);
    const currentFullRows = currentFullMonth.filter(
      (s) => managerOfSale(s) === manager,
    );
    const prevMonthRows = prevMonth.filter((s) => managerOfSale(s) === manager);
    const prevYearRows = prevYear.filter((s) => managerOfSale(s) === manager);
    const managerEst = ests
      .filter((e) => e.month === month && managerOfEst(e) === manager)
      .reduce((total, row) => total + Number(row.amount || 0), 0);
    const currentSales = sum(currentRows, "salesAmount");
    const prevMonthSales = sum(prevMonthRows, "salesAmount");
    const prevYearSales = sum(prevYearRows, "salesAmount");
    const profitAmount = sum(currentRows, "profitAmount");

    return {
      manager,
      prevYearSales,
      prevMonthSales,
      currentSales,
      fullMonthSales: sum(currentFullRows, "salesAmount"),
      prevYearRate: prevYearSales
        ? ((currentSales - prevYearSales) / prevYearSales) * 100
        : 0,
      prevMonthRate: prevMonthSales
        ? ((currentSales - prevMonthSales) / prevMonthSales) * 100
        : 0,
      est: managerEst,
      estRate: managerEst ? (currentSales / managerEst) * 100 : 0,
      profitAmount,
      profitRate: weightedProfitRate(currentRows),
    };
  });

  const total = rows.reduce(
    (acc, row) => ({
      prevYearSales: acc.prevYearSales + row.prevYearSales,
      prevMonthSales: acc.prevMonthSales + row.prevMonthSales,
      currentSales: acc.currentSales + row.currentSales,
      fullMonthSales: acc.fullMonthSales + row.fullMonthSales,
      est: acc.est + row.est,
      profitAmount: acc.profitAmount + row.profitAmount,
    }),
    {
      prevYearSales: 0,
      prevMonthSales: 0,
      currentSales: 0,
      fullMonthSales: 0,
      est: 0,
      profitAmount: 0,
    },
  );

  const dashboardManagerExcelRows = rows.map((r) => ({
    담당자: r.manager,
    전년동월: r.prevYearSales,
    전년대비: pct(r.prevYearRate),
    전월: r.prevMonthSales,
    전월대비: pct(r.prevMonthRate),
    당일까지매출: r.currentSales,
    당월전체매출: r.fullMonthSales,
    EST: r.est,
    EST달성률: pct(r.estRate),
    이익금액: r.profitAmount,
    이익률: pct(r.profitRate),
  }));

  return (
    <div className="overflow-hidden rounded-2xl border border-slate-300 bg-white shadow-sm">
      <div className="flex flex-wrap items-center justify-between gap-3 border-b border-slate-300 px-4 py-3">
        <div>
          <div className="text-base font-extrabold text-slate-900">
            담당자별 매출 요약
          </div>
        </div>
        <button
          type="button"
          onClick={() =>
            exportExcel(dashboardManagerExcelRows, `담당자별_매출요약_${month}`)
          }
          className="rounded-lg bg-emerald-600 px-3 py-1.5 text-xs font-semibold text-white shadow-sm hover:bg-emerald-700"
        >
          엑셀 다운로드
        </button>
      </div>
      <div className="overflow-auto">
        <table className="w-full min-w-[1100px] border-separate border-spacing-0 text-center text-[12px] whitespace-nowrap">
          <thead>
            <tr className="bg-slate-100">
              <th className="sticky top-0 z-10 border border-slate-300 bg-slate-100 px-3 py-2 font-bold text-slate-700">
                담당자
              </th>
              <th className="sticky top-0 z-10 border border-slate-300 bg-[#F7FCEB] px-3 py-2 font-bold text-black">
                전년동월
              </th>
              <th className="sticky top-0 z-10 border border-slate-300 bg-[#F7FCEB] px-3 py-2 font-bold text-black">
                전년대비
              </th>
              <th className="sticky top-0 z-10 border border-slate-300 bg-[#F3FAFD] px-3 py-2 font-bold text-black">
                전월
              </th>
              <th className="sticky top-0 z-10 border border-slate-300 bg-[#F3FAFD] px-3 py-2 font-bold text-black">
                전월대비
              </th>
              <th className="sticky top-0 z-10 border border-slate-300 bg-[#FFF7FA] px-3 py-2 font-bold text-black">
                당일까지 매출
              </th>
              <th className="sticky top-0 z-10 border border-slate-300 bg-[#FFF7FA] px-3 py-2 font-bold text-black">
                당월 전체 매출
              </th>
              <th className="sticky top-0 z-10 border border-slate-300 bg-[#FFFDF2] px-3 py-2 font-bold text-black">
                EST
              </th>
              <th className="sticky top-0 z-10 border border-slate-300 bg-[#FFFDF2] px-3 py-2 font-bold text-black">
                EST 달성률
              </th>
              <th className="sticky top-0 z-10 border border-slate-300 bg-[#FFF9F3] px-3 py-2 font-bold text-black">
                이익금액
              </th>
              <th className="sticky top-0 z-10 border border-slate-300 bg-[#FFF9F3] px-3 py-2 font-bold text-black">
                이익률
              </th>
            </tr>
          </thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r.manager} className="hover:bg-slate-50">
                <td className="border border-slate-300 px-3 py-2 text-center text-sm font-extrabold text-slate-900">
                  {r.manager}
                </td>
                <td className="border border-slate-300 px-3 py-2 text-right font-semibold">
                  {won(r.prevYearSales)}
                </td>
                <td
                  className={`border border-slate-300 px-3 py-2 text-right font-semibold ${r.prevYearRate >= 0 ? "text-emerald-700" : "text-red-600"}`}
                >
                  {pct(r.prevYearRate)}
                </td>
                <td className="border border-slate-300 px-3 py-2 text-right font-semibold">
                  {won(r.prevMonthSales)}
                </td>
                <td
                  className={`border border-slate-300 px-3 py-2 text-right font-semibold ${r.prevMonthRate >= 0 ? "text-emerald-700" : "text-red-600"}`}
                >
                  {pct(r.prevMonthRate)}
                </td>
                <td className="border border-slate-300 px-3 py-2 text-right text-blue-700 font-extrabold">
                  {won(r.currentSales)}
                </td>
                <td className="border border-slate-300 px-3 py-2 text-right font-semibold">
                  {won(r.fullMonthSales)}
                </td>
                <td className="border border-slate-300 px-3 py-2 text-right font-semibold">
                  {won(r.est)}
                </td>
                <td className="border border-slate-300 px-3 py-2 text-right font-semibold">
                  {pct(r.estRate)}
                </td>
                <td className="border border-slate-300 px-3 py-2 text-right font-semibold">
                  {won(r.profitAmount)}
                </td>
                <td className="border border-slate-300 px-3 py-2 text-right font-semibold">
                  {pct(r.profitRate)}
                </td>
              </tr>
            ))}
            <tr className="bg-slate-100 font-extrabold">
              <td className="border border-slate-300 px-3 py-2 text-center">
                합계
              </td>
              <td className="border border-slate-300 px-3 py-2 text-right">
                {won(total.prevYearSales)}
              </td>
              <td className="border border-slate-300 px-3 py-2 text-right">
                {pct(
                  total.prevYearSales
                    ? ((total.currentSales - total.prevYearSales) /
                        total.prevYearSales) *
                        100
                    : 0,
                )}
              </td>
              <td className="border border-slate-300 px-3 py-2 text-right">
                {won(total.prevMonthSales)}
              </td>
              <td className="border border-slate-300 px-3 py-2 text-right">
                {pct(
                  total.prevMonthSales
                    ? ((total.currentSales - total.prevMonthSales) /
                        total.prevMonthSales) *
                        100
                    : 0,
                )}
              </td>
              <td className="border border-slate-300 px-3 py-2 text-right text-blue-700">
                {won(total.currentSales)}
              </td>
              <td className="border border-slate-300 px-3 py-2 text-right">
                {won(total.fullMonthSales)}
              </td>
              <td className="border border-slate-300 px-3 py-2 text-right">
                {won(total.est)}
              </td>
              <td className="border border-slate-300 px-3 py-2 text-right">
                {pct(total.est ? (total.currentSales / total.est) * 100 : 0)}
              </td>
              <td className="border border-slate-300 px-3 py-2 text-right">
                {won(total.profitAmount)}
              </td>
              <td className="border border-slate-300 px-3 py-2 text-right">
                {pct(
                  total.currentSales
                    ? (total.profitAmount / total.currentSales) * 100
                    : 0,
                )}
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}

type ItemMetric = {
  qty: number;
  sales: number;
  cost: number;
  profit: number;
  // 업로드한 로우파일의 이익률을 매출 가중치로 집계합니다.
  profitRateWeighted: number;
  profitRateSales: number;
};

type ItemDetailSortKey =
  | "saleDate"
  | "storeName"
  | "itemCode"
  | "itemName"
  | "quantity"
  | "salesAmount"
  | "costAmount"
  | "profitAmount"
  | "profitRate"
  | "lastOrderDate";

type ItemBrandSortKey =
  | "brand"
  | "current"
  | "prevMonth"
  | "prevMonthRate"
  | "prevYear"
  | "prevYearRate"
  | "storeCount";
type ItemStoreSortKey =
  | "brand"
  | "code"
  | "name"
  | "manager"
  | "current"
  | "prevMonthRate"
  | "prevYearRate";
type ItemRowSortKey =
  | "itemCode"
  | "itemName"
  | "currentQty"
  | "prevMonthQty"
  | "prevMonthDiff"
  | "prevMonthQtyRate"
  | "prevYearQty"
  | "prevYearDiff"
  | "prevYearQtyRate"
  | "currentSales"
  | "prevMonthSales"
  | "prevMonthSalesDiff"
  | "prevMonthSalesRate"
  | "prevYearSales"
  | "prevYearSalesDiff"
  | "prevYearSalesRate";

type ItemAnalysisRow = {
  itemCode: string;
  itemName: string;
  current: ItemMetric;
  prevMonth: ItemMetric;
  prevYear: ItemMetric;
};

function emptyItemMetric(): ItemMetric {
  return {
    qty: 0,
    sales: 0,
    cost: 0,
    profit: 0,
    profitRateWeighted: 0,
    profitRateSales: 0,
  };
}

function addItemMetric(metric: ItemMetric, row: SalesRecord) {
  const salesAmount = Number(row.salesAmount || 0);
  const uploadedProfitRate = Number(row.profitRate || 0);
  metric.qty += Number(row.quantity || 0);
  metric.sales += salesAmount;
  metric.cost += Number(row.costAmount || 0);
  metric.profit += Number(row.profitAmount || 0);
  if (Number.isFinite(uploadedProfitRate) && salesAmount !== 0) {
    metric.profitRateWeighted += uploadedProfitRate * salesAmount;
    metric.profitRateSales += salesAmount;
  }
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
  // 행과 Subtotal 모두 업로드 로우파일의 이익률을 기본값으로 사용합니다.
  if (metric.profitRateSales) {
    return metric.profitRateWeighted / metric.profitRateSales;
  }
  return 0;
}

function itemStoreMatches(store: Store, search: string) {
  const q = search.trim().toLowerCase();
  if (!q) return true;
  return [
    store.brand,
    store.code,
    store.name,
    store.channel,
    store.storeType,
    store.manager,
  ].some((v) =>
    String(v || "")
      .toLowerCase()
      .includes(q),
  );
}

function salesInRangeByStore(
  sales: SalesRecord[],
  storeCode: string,
  start: string,
  end: string,
) {
  return sales.filter(
    (r) => r.storeCode === storeCode && inRange(r.saleDate, start, end),
  );
}

function sumSalesInRangeByStore(
  sales: SalesRecord[],
  storeCode: string,
  start: string,
  end: string,
) {
  return salesInRangeByStore(sales, storeCode, start, end).reduce(
    (total, r) => total + Number(r.salesAmount || 0),
    0,
  );
}

function buildItemAnalysisRows(
  sales: SalesRecord[],
  storeCode: string,
  currentStart: string,
  currentEnd: string,
  prevStart: string,
  prevEnd: string,
  prevYearStart: string,
  prevYearEnd: string,
): ItemAnalysisRow[] {
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
    if (inRange(r.saleDate, currentStart, currentEnd))
      addItemMetric(row.current, r);
    if (inRange(r.saleDate, prevStart, prevEnd))
      addItemMetric(row.prevMonth, r);
    if (inRange(r.saleDate, prevYearStart, prevYearEnd))
      addItemMetric(row.prevYear, r);
  });

  return Array.from(itemMap.values())
    .filter(
      (r) =>
        r.current.qty ||
        r.prevMonth.qty ||
        r.prevYear.qty ||
        r.current.sales ||
        r.prevMonth.sales ||
        r.prevYear.sales,
    )
    .sort((a, b) => b.current.sales - a.current.sales);
}

function PopupTh({
  children,
  right = false,
  rowSpan,
  colSpan,
  top = "top-0",
  className = "bg-white",
}: {
  children: React.ReactNode;
  right?: boolean;
  rowSpan?: number;
  colSpan?: number;
  top?: string;
  className?: string;
}) {
  return (
    <th
      rowSpan={rowSpan}
      colSpan={colSpan}
      className={`sticky ${top} z-50 border border-gray-300 px-3 py-2 font-bold shadow-[0_2px_0_0_#e2e8f0] ${className} text-center`}
    >
      {children}
    </th>
  );
}

function ItemAnalysisSortableTh<K extends string>({
  children,
  sortKey,
  sortConfig,
  onSort,
  right = false,
  rowSpan,
  top = "top-0",
  className = "bg-white",
}: {
  children: React.ReactNode;
  sortKey: K;
  sortConfig: { key: K; direction: SortDirection };
  onSort: (key: K) => void;
  right?: boolean;
  rowSpan?: number;
  top?: string;
  className?: string;
}) {
  return (
    <PopupTh right={right} rowSpan={rowSpan} top={top} className={className}>
      <button
        type="button"
        onClick={() => onSort(sortKey)}
        className="flex w-full items-center justify-center gap-1 text-center"
        title="정렬"
      >
        <span>{children}</span>
        <span className="text-[11px] text-black">
          {sortArrow(sortConfig.key === sortKey, sortConfig.direction)}
        </span>
      </button>
    </PopupTh>
  );
}

function ItemAnalysis({
  stores,
  sales,
  month,
  date,
  pageTitle = "품목분석",
}: {
  stores: Store[];
  sales: SalesRecord[];
  month: string;
  date: string;
  pageTitle?: string;
}) {
  const [mode, setMode] = useState<"브랜드별" | "매장별">("브랜드별");
  const [searchDraft, setSearchDraft] = useState("");
  const [search, setSearch] = useState("");
  const [selectedBrand, setSelectedBrand] = useState("");
  const [selectedStoreCode, setSelectedStoreCode] = useState("");
  const [selectedItemCode, setSelectedItemCode] = useState("");
  const [brandSortConfig, setBrandSortConfig] = useState<{
    key: ItemBrandSortKey;
    direction: SortDirection;
  }>({ key: "current", direction: "desc" });
  const [storeSortConfig, setStoreSortConfig] = useState<{
    key: ItemStoreSortKey;
    direction: SortDirection;
  }>({ key: "current", direction: "desc" });
  const [itemSortConfig, setItemSortConfig] = useState<{
    key: ItemRowSortKey;
    direction: SortDirection;
  }>({ key: "currentSales", direction: "desc" });
  const [detailSortConfig, setDetailSortConfig] = useState<{
    key: ItemDetailSortKey;
    direction: SortDirection;
  }>({ key: "saleDate", direction: "asc" });
  const defaultCurrentEnd = date.startsWith(month) ? date : monthEnd(month);
  const [analysisStart, setAnalysisStart] = useState(monthStart(month));
  const [analysisEnd, setAnalysisEnd] = useState(defaultCurrentEnd);
  const [prevAnalysisStart, setPrevAnalysisStart] = useState(sameDayPrevMonth(monthStart(month)));
  const [prevAnalysisEnd, setPrevAnalysisEnd] = useState(sameDayPrevMonth(defaultCurrentEnd));
  const [prevYearAnalysisStart, setPrevYearAnalysisStart] = useState(sameDayPrevYear(monthStart(month)));
  const [prevYearAnalysisEnd, setPrevYearAnalysisEnd] = useState(sameDayPrevYear(defaultCurrentEnd));
  const [appliedPeriods, setAppliedPeriods] = useState(() => ({
    currentStart: monthStart(month),
    currentEnd: defaultCurrentEnd,
    prevStart: sameDayPrevMonth(monthStart(month)),
    prevEnd: sameDayPrevMonth(defaultCurrentEnd),
    prevYearStart: sameDayPrevYear(monthStart(month)),
    prevYearEnd: sameDayPrevYear(defaultCurrentEnd),
  }));

  useEffect(() => {
    const nextCurrentStart = monthStart(month);
    const nextCurrentEnd = date.startsWith(month) ? date : monthEnd(month);
    const nextPrevStart = sameDayPrevMonth(nextCurrentStart);
    const nextPrevEnd = sameDayPrevMonth(nextCurrentEnd);
    const nextPrevYearStart = sameDayPrevYear(nextCurrentStart);
    const nextPrevYearEnd = sameDayPrevYear(nextCurrentEnd);
    setAnalysisStart(nextCurrentStart);
    setAnalysisEnd(nextCurrentEnd);
    setPrevAnalysisStart(nextPrevStart);
    setPrevAnalysisEnd(nextPrevEnd);
    setPrevYearAnalysisStart(nextPrevYearStart);
    setPrevYearAnalysisEnd(nextPrevYearEnd);
    setAppliedPeriods({
      currentStart: nextCurrentStart, currentEnd: nextCurrentEnd,
      prevStart: nextPrevStart, prevEnd: nextPrevEnd,
      prevYearStart: nextPrevYearStart, prevYearEnd: nextPrevYearEnd,
    });
    setSelectedBrand("");
    setSelectedStoreCode("");
    setSelectedItemCode("");
  }, [month, date]);

  const normalizePeriod = (start: string, end: string) =>
    start <= end ? [start, end] as const : [end, start] as const;
  const [currentStart, currentEnd] = normalizePeriod(appliedPeriods.currentStart, appliedPeriods.currentEnd);
  const [prevStart, prevEnd] = normalizePeriod(appliedPeriods.prevStart, appliedPeriods.prevEnd);
  const [prevYearStart, prevYearEnd] = normalizePeriod(appliedPeriods.prevYearStart, appliedPeriods.prevYearEnd);

  function applyPeriodSearch() {
    const [nextCurrentStart, nextCurrentEnd] = normalizePeriod(analysisStart, analysisEnd);
    const [nextPrevStart, nextPrevEnd] = normalizePeriod(prevAnalysisStart, prevAnalysisEnd);
    const [nextPrevYearStart, nextPrevYearEnd] = normalizePeriod(prevYearAnalysisStart, prevYearAnalysisEnd);
    setAppliedPeriods({
      currentStart: nextCurrentStart, currentEnd: nextCurrentEnd,
      prevStart: nextPrevStart, prevEnd: nextPrevEnd,
      prevYearStart: nextPrevYearStart, prevYearEnd: nextPrevYearEnd,
    });
    setSelectedBrand("");
    setSelectedStoreCode("");
    setSelectedItemCode("");
  }
  const normalizedSearch = search.trim().toLowerCase();
  const storeByCode = useMemo(
    () => new Map(stores.map((s) => [s.code, s])),
    [stores],
  );
  const activeStores = useMemo(
    () =>
      stores.filter(
        (s) => s.status !== "거래종료" && itemStoreMatches(s, normalizedSearch),
      ),
    [stores, normalizedSearch],
  );
  const selectedStore = storeByCode.get(selectedStoreCode);

  const brandRows = useMemo(() => {
    const map = new Map<
      string,
      {
        brand: string;
        stores: Store[];
        current: number;
        prevMonth: number;
        prevYear: number;
      }
    >();
    activeStores.forEach((s) => {
      const brand = displayBrand(s.brand);
      if (!map.has(brand))
        map.set(brand, {
          brand,
          stores: [],
          current: 0,
          prevMonth: 0,
          prevYear: 0,
        });
      const row = map.get(brand)!;
      row.stores.push(s);
      row.current += sumSalesInRangeByStore(
        sales,
        s.code,
        currentStart,
        currentEnd,
      );
      row.prevMonth += sumSalesInRangeByStore(
        sales,
        s.code,
        prevStart,
        prevEnd,
      );
      row.prevYear += sumSalesInRangeByStore(
        sales,
        s.code,
        prevYearStart,
        prevYearEnd,
      );
    });
    return Array.from(map.values()).sort((a, b) => b.current - a.current);
  }, [
    activeStores,
    sales,
    currentStart,
    currentEnd,
    prevStart,
    prevEnd,
    prevYearStart,
    prevYearEnd,
  ]);

  const storeRows = useMemo(() => {
    const base =
      mode === "브랜드별" && selectedBrand
        ? activeStores.filter((s) => displayBrand(s.brand) === selectedBrand)
        : activeStores;
    return base
      .map((s) => ({
        store: s,
        current: sumSalesInRangeByStore(
          sales,
          s.code,
          currentStart,
          currentEnd,
        ),
        prevMonth: sumSalesInRangeByStore(sales, s.code, prevStart, prevEnd),
        prevYear: sumSalesInRangeByStore(
          sales,
          s.code,
          prevYearStart,
          prevYearEnd,
        ),
      }))
      .sort((a, b) => b.current - a.current);
  }, [
    activeStores,
    selectedBrand,
    mode,
    sales,
    currentStart,
    currentEnd,
    prevStart,
    prevEnd,
    prevYearStart,
    prevYearEnd,
  ]);

  const itemRows = useMemo(
    () =>
      selectedStoreCode
        ? buildItemAnalysisRows(
            sales,
            selectedStoreCode,
            currentStart,
            currentEnd,
            prevStart,
            prevEnd,
            prevYearStart,
            prevYearEnd,
          )
        : [],
    [
      sales,
      selectedStoreCode,
      currentStart,
      currentEnd,
      prevStart,
      prevEnd,
      prevYearStart,
      prevYearEnd,
    ],
  );

  const itemSalesSubtotal = useMemo(
    () => ({
      currentSales: itemRows.reduce(
        (total, row) => total + Number(row.current.sales || 0),
        0,
      ),
      prevMonthSales: itemRows.reduce(
        (total, row) => total + Number(row.prevMonth.sales || 0),
        0,
      ),
      prevYearSales: itemRows.reduce(
        (total, row) => total + Number(row.prevYear.sales || 0),
        0,
      ),
    }),
    [itemRows],
  );

  const sortedBrandRows = useMemo(() => {
    const valueOf = (r: (typeof brandRows)[number]) => {
      if (brandSortConfig.key === "brand") return r.brand;
      if (brandSortConfig.key === "current") return r.current;
      if (brandSortConfig.key === "prevMonth") return r.prevMonth;
      if (brandSortConfig.key === "prevMonthRate")
        return itemMetricRate(r.current, r.prevMonth);
      if (brandSortConfig.key === "prevYear") return r.prevYear;
      if (brandSortConfig.key === "prevYearRate")
        return itemMetricRate(r.current, r.prevYear);
      return r.stores.length;
    };
    return [...brandRows].sort((a, b) => {
      const aValue = valueOf(a);
      const bValue = valueOf(b);
      const result =
        typeof aValue === "string" || typeof bValue === "string"
          ? String(aValue ?? "").localeCompare(String(bValue ?? ""), "ko-KR", {
              numeric: true,
            })
          : Number(aValue || 0) - Number(bValue || 0);
      return brandSortConfig.direction === "asc" ? result : -result;
    });
  }, [brandRows, brandSortConfig]);

  const sortedStoreRows = useMemo(() => {
    const valueOf = (r: (typeof storeRows)[number]) => {
      if (storeSortConfig.key === "brand") return r.store.brand;
      if (storeSortConfig.key === "code") return r.store.code;
      if (storeSortConfig.key === "name") return r.store.name;
      if (storeSortConfig.key === "manager") return r.store.manager || "미지정";
      if (storeSortConfig.key === "current") return r.current;
      if (storeSortConfig.key === "prevMonthRate")
        return itemMetricRate(r.current, r.prevMonth);
      return itemMetricRate(r.current, r.prevYear);
    };
    return [...storeRows].sort((a, b) => {
      const aValue = valueOf(a);
      const bValue = valueOf(b);
      const result =
        typeof aValue === "string" || typeof bValue === "string"
          ? String(aValue ?? "").localeCompare(String(bValue ?? ""), "ko-KR", {
              numeric: true,
            })
          : Number(aValue || 0) - Number(bValue || 0);
      return storeSortConfig.direction === "asc" ? result : -result;
    });
  }, [storeRows, storeSortConfig]);

  const sortedItemRows = useMemo(() => {
    const valueOf = (r: ItemAnalysisRow) => {
      if (itemSortConfig.key === "itemCode") return r.itemCode;
      if (itemSortConfig.key === "itemName") return r.itemName;
      if (itemSortConfig.key === "currentQty") return r.current.qty;
      if (itemSortConfig.key === "prevMonthQty") return r.prevMonth.qty;
      if (itemSortConfig.key === "prevMonthDiff")
        return itemMetricDiff(r.current.qty, r.prevMonth.qty);
      if (itemSortConfig.key === "prevMonthQtyRate")
        return itemMetricRate(r.current.qty, r.prevMonth.qty);
      if (itemSortConfig.key === "prevYearQty") return r.prevYear.qty;
      if (itemSortConfig.key === "prevYearDiff")
        return itemMetricDiff(r.current.qty, r.prevYear.qty);
      if (itemSortConfig.key === "prevYearQtyRate")
        return itemMetricRate(r.current.qty, r.prevYear.qty);
      if (itemSortConfig.key === "currentSales") return r.current.sales;
      if (itemSortConfig.key === "prevMonthSales") return r.prevMonth.sales;
      if (itemSortConfig.key === "prevMonthSalesDiff")
        return itemMetricDiff(r.current.sales, r.prevMonth.sales);
      if (itemSortConfig.key === "prevMonthSalesRate")
        return itemMetricRate(r.current.sales, r.prevMonth.sales);
      if (itemSortConfig.key === "prevYearSales") return r.prevYear.sales;
      if (itemSortConfig.key === "prevYearSalesDiff")
        return itemMetricDiff(r.current.sales, r.prevYear.sales);
      return itemMetricRate(r.current.sales, r.prevYear.sales);
    };
    return [...itemRows].sort((a, b) => {
      const aValue = valueOf(a);
      const bValue = valueOf(b);
      const result =
        typeof aValue === "string" || typeof bValue === "string"
          ? String(aValue ?? "").localeCompare(String(bValue ?? ""), "ko-KR", {
              numeric: true,
            })
          : Number(aValue || 0) - Number(bValue || 0);
      return itemSortConfig.direction === "asc" ? result : -result;
    });
  }, [itemRows, itemSortConfig]);
  const selectedItem = itemRows.find((r) => r.itemCode === selectedItemCode);
  const detailRows = useMemo(() => {
    if (!selectedStoreCode || !selectedItemCode) return [];
    return sales.filter(
      (r) =>
        r.storeCode === selectedStoreCode &&
        r.itemCode === selectedItemCode &&
        inRange(r.saleDate, currentStart, currentEnd),
    );
  }, [sales, selectedStoreCode, selectedItemCode, currentStart, currentEnd]);

  const sortedDetailRows = useMemo(() => {
    return [...detailRows].sort((a, b) => {
      const aValue = a[detailSortConfig.key];
      const bValue = b[detailSortConfig.key];
      let result = 0;
      if (typeof aValue === "string" || typeof bValue === "string") {
        result = String(aValue ?? "").localeCompare(
          String(bValue ?? ""),
          "ko-KR",
          { numeric: true },
        );
      } else {
        result = Number(aValue || 0) - Number(bValue || 0);
      }
      return detailSortConfig.direction === "asc" ? result : -result;
    });
  }, [detailRows, detailSortConfig]);

  const requestBrandSort = (key: ItemBrandSortKey) => {
    setBrandSortConfig((prev) => ({
      key,
      direction: prev.key === key && prev.direction === "desc" ? "asc" : "desc",
    }));
  };

  const requestStoreSort = (key: ItemStoreSortKey) => {
    setStoreSortConfig((prev) => ({
      key,
      direction: prev.key === key && prev.direction === "desc" ? "asc" : "desc",
    }));
  };

  const requestItemSort = (key: ItemRowSortKey) => {
    setItemSortConfig((prev) => ({
      key,
      direction: prev.key === key && prev.direction === "desc" ? "asc" : "desc",
    }));
  };

  const requestDetailSort = (key: ItemDetailSortKey) => {
    setDetailSortConfig((prev) => ({
      key,
      direction: prev.key === key && prev.direction === "desc" ? "asc" : "desc",
    }));
  };

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
            <div className="text-base font-bold text-slate-900">{pageTitle}</div>
            <div className="mt-1 text-sm text-slate-500">
              현재 {currentStart} ~ {currentEnd} / 전월 {prevStart} ~ {prevEnd}{" "}
              / 전년 {prevYearStart} ~ {prevYearEnd}
            </div>
          </div>
          <div className="flex flex-wrap items-center gap-2 text-xs">
            <input
              value={searchDraft}
              onChange={(e) => {
                const next = e.target.value;
                setSearchDraft(next);
                if (!next.trim()) setSearch("");
              }}
              onKeyDown={(e) => {
                if (e.key === "Enter") applySearch();
              }}
              placeholder="브랜드/거래처/품목/담당자 검색 후 Enter"
              className="h-8 w-[260px] rounded-lg border border-slate-300 bg-white px-3 text-xs outline-none focus:border-blue-500"
            />
            <button
              type="button"
              onClick={applySearch}
              className="h-8 rounded-lg bg-slate-800 px-3 text-sm font-semibold text-white hover:bg-slate-700"
            >
              검색
            </button>
            <div className="flex flex-wrap items-center gap-2 rounded-xl border border-slate-200 bg-slate-50 px-3 py-2">
              <span className="font-bold text-slate-700">당월</span>
              <input type="date" value={analysisStart} onChange={(e) => setAnalysisStart(e.target.value)} className="h-8 rounded-lg border border-slate-300 px-2 text-xs" />
              <span>~</span>
              <input type="date" value={analysisEnd} onChange={(e) => setAnalysisEnd(e.target.value)} className="h-8 rounded-lg border border-slate-300 px-2 text-xs" />
              <span className="ml-1 font-bold text-slate-700">전월</span>
              <input type="date" value={prevAnalysisStart} onChange={(e) => setPrevAnalysisStart(e.target.value)} className="h-8 rounded-lg border border-slate-300 px-2 text-xs" />
              <span>~</span>
              <input type="date" value={prevAnalysisEnd} onChange={(e) => setPrevAnalysisEnd(e.target.value)} className="h-8 rounded-lg border border-slate-300 px-2 text-xs" />
              <span className="ml-1 font-bold text-slate-700">전년동월</span>
              <input type="date" value={prevYearAnalysisStart} onChange={(e) => setPrevYearAnalysisStart(e.target.value)} className="h-8 rounded-lg border border-slate-300 px-2 text-xs" />
              <span>~</span>
              <input type="date" value={prevYearAnalysisEnd} onChange={(e) => setPrevYearAnalysisEnd(e.target.value)} className="h-8 rounded-lg border border-slate-300 px-2 text-xs" />
              <button type="button" onClick={applyPeriodSearch} className="h-8 rounded-lg bg-blue-600 px-4 text-xs font-bold text-white hover:bg-blue-700">조회</button>
            </div>
          </div>
          <div className="flex rounded-xl bg-slate-100 p-1 text-sm font-semibold">
            {(["브랜드별", "매장별"] as const).map((m) => (
              <button
                key={m}
                onClick={() => {
                  setMode(m);
                  backToTop();
                }}
                className={`rounded-lg px-4 py-2 ${mode === m ? "bg-white text-blue-700 shadow-sm" : "text-slate-600 hover:text-blue-700"}`}
              >
                {m}
              </button>
            ))}
          </div>
        </div>
      </div>

      {!selectedStoreCode && mode === "브랜드별" && !selectedBrand && (
        <div className="overflow-hidden rounded-2xl border border-slate-300 bg-white shadow-sm">
          <div className="border-b border-slate-300 px-4 py-3 text-sm font-bold text-slate-800">
            브랜드별 요약
          </div>
          <div className="max-h-[74vh] overflow-auto isolate">
            <table className="w-full min-w-[1100px] connected-two-tier border-separate border-spacing-0 text-center text-[15px] leading-none whitespace-nowrap">
              <thead>
                <tr>
                  <ItemAnalysisSortableTh rowSpan={2} sortKey="brand" sortConfig={brandSortConfig} onSort={requestBrandSort}>브랜드</ItemAnalysisSortableTh>
                  <PopupTh colSpan={2} className="period-group-start bg-[#F7FCEB] text-center text-black">전년동월</PopupTh>
                  <PopupTh colSpan={2} className="period-group-start bg-[#F3FAFD] text-center text-black">전월</PopupTh>
                  <PopupTh colSpan={1} className="period-group-start bg-[#FFF7FA] text-center text-black">당월</PopupTh>
                  <ItemAnalysisSortableTh rowSpan={2} sortKey="storeCount" sortConfig={brandSortConfig} onSort={requestBrandSort} right>거래처수</ItemAnalysisSortableTh>
                  <PopupTh rowSpan={2}>상세</PopupTh>
                </tr>
                <tr>
                  <ItemAnalysisSortableTh top="top-[37px]" className="period-group-start bg-[#F7FCEB] text-black" sortKey="prevYear" sortConfig={brandSortConfig} onSort={requestBrandSort} right>매출</ItemAnalysisSortableTh>
                  <ItemAnalysisSortableTh top="top-[37px]" className="bg-[#F7FCEB] text-black" sortKey="prevYearRate" sortConfig={brandSortConfig} onSort={requestBrandSort} right>당월 대비</ItemAnalysisSortableTh>
                  <ItemAnalysisSortableTh top="top-[37px]" className="period-group-start bg-[#F3FAFD] text-black" sortKey="prevMonth" sortConfig={brandSortConfig} onSort={requestBrandSort} right>매출</ItemAnalysisSortableTh>
                  <ItemAnalysisSortableTh top="top-[37px]" className="bg-[#F3FAFD] text-black" sortKey="prevMonthRate" sortConfig={brandSortConfig} onSort={requestBrandSort} right>당월 대비</ItemAnalysisSortableTh>
                  <ItemAnalysisSortableTh top="top-[37px]" className="period-group-start bg-[#FFF7FA] text-black" sortKey="current" sortConfig={brandSortConfig} onSort={requestBrandSort} right>매출</ItemAnalysisSortableTh>
                </tr>
              </thead>
              <tbody>
                {sortedBrandRows.map((r) => (
                  <tr key={r.brand} className="hover:bg-blue-50">
                    <td className="border border-slate-300 p-2 font-semibold">
                      {r.brand}
                    </td>
                    <td className="border border-slate-300 p-2 text-right">{won(r.prevYear)}</td>
                    <td className={`border border-slate-300 p-2 text-right ${itemMetricDiff(r.current, r.prevYear) >= 0 ? "text-emerald-600" : "text-red-600"}`}>{itemSignedPct(itemMetricRate(r.current, r.prevYear))}</td>
                    <td className="border border-slate-300 p-2 text-right">{won(r.prevMonth)}</td>
                    <td className={`border border-slate-300 p-2 text-right ${itemMetricDiff(r.current, r.prevMonth) >= 0 ? "text-emerald-600" : "text-red-600"}`}>{itemSignedPct(itemMetricRate(r.current, r.prevMonth))}</td>
                    <td className="border border-slate-300 p-2 text-right font-bold text-blue-700">{won(r.current)}</td>
                    <td className="border border-slate-300 p-2 text-right">
                      {r.stores.length}
                    </td>
                    <td className="border border-slate-300 p-2">
                      <button
                        onClick={() => setSelectedBrand(r.brand)}
                        className="rounded-lg border border-slate-300 bg-white px-3 py-1 text-sm font-semibold text-black hover:border-blue-300 hover:bg-blue-50"
                      >
                        거래처 보기
                      </button>
                    </td>
                  </tr>
                ))}
                {!brandRows.length && (
                  <tr>
                    <td
                      colSpan={8}
                      className="border border-slate-300 p-8 text-center text-slate-500"
                    >
                      표시할 데이터가 없습니다.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {!selectedStoreCode && (mode === "매장별" || selectedBrand) && (
        <div className="overflow-hidden rounded-2xl border border-slate-300 bg-white shadow-sm">
          <div className="flex items-center justify-between border-b border-slate-300 px-4 py-3">
            <div className="text-sm font-bold text-slate-800">
              {selectedBrand ? `${selectedBrand} 거래처 목록` : "거래처별 요약"}
            </div>
            {selectedBrand && (
              <button
                onClick={backToTop}
                className="rounded-lg bg-slate-100 px-3 py-1 text-sm font-semibold text-slate-700 hover:bg-slate-200"
              >
                ← 브랜드 목록
              </button>
            )}
          </div>
          <div className="max-h-[74vh] overflow-auto isolate">
            <table className="w-full min-w-[900px] connected-two-tier border-separate border-spacing-0 text-center text-[12px] text-slate-900 whitespace-nowrap">
              {mode === "매장별" && !selectedBrand ? (
                <>
                  <thead>
                    <tr>
                      <ItemAnalysisSortableTh rowSpan={2} sortKey="name" sortConfig={storeSortConfig} onSort={requestStoreSort}>거래처명</ItemAnalysisSortableTh>
                      <ItemAnalysisSortableTh rowSpan={2} sortKey="manager" sortConfig={storeSortConfig} onSort={requestStoreSort}>담당자</ItemAnalysisSortableTh>
                      <PopupTh rowSpan={2} className="bg-slate-100 text-center font-bold text-slate-900">구분</PopupTh>
                      <PopupTh colSpan={3} className="period-group-start bg-white text-center text-[13px] font-extrabold text-black">Time Gone 대비</PopupTh>
                      <PopupTh rowSpan={2} className="bg-slate-100 text-center font-bold text-slate-900">상세</PopupTh>
                    </tr>
                    <tr>
                      <PopupTh top="top-[37px]" className="period-group-start bg-[#F7FCEB] text-center font-bold text-black">전년동월</PopupTh>
                      <PopupTh top="top-[37px]" className="period-subgroup-start bg-[#F3FAFD] text-center font-bold text-black">전월</PopupTh>
                      <ItemAnalysisSortableTh top="top-[37px]" className="period-group-start bg-[#FFF7FA] text-black" sortKey="current" sortConfig={storeSortConfig} onSort={requestStoreSort} right>당일까지 매출</ItemAnalysisSortableTh>
                    </tr>
                  </thead>
                  <tbody>
                    {sortedStoreRows.map((r) => (
                      <tr key={r.store.code} className="hover:bg-slate-50">
                        <td className="border border-slate-300 px-3 py-2 text-center text-[12px] font-semibold text-slate-900 whitespace-nowrap">
                          {r.store.name}
                        </td>
                        <td className="border border-slate-300 px-3 py-2 text-center text-[12px] font-semibold text-slate-900 whitespace-nowrap">
                          {r.store.manager || "미지정"}
                        </td>
                        <td className="border border-slate-300 px-3 py-2 text-center text-[12px] font-semibold text-slate-900 whitespace-nowrap">
                          {r.store.storeType === "매장" ? "매장" : "비매장"}
                        </td>
                        <td className="border border-slate-300 px-3 py-2 text-right text-[13px] font-bold text-slate-900 tabular-nums whitespace-nowrap">{won(r.prevYear)}</td>
                        <td className="border border-slate-300 px-3 py-2 text-right text-[13px] font-bold text-slate-900 tabular-nums whitespace-nowrap">{won(r.prevMonth)}</td>
                        <td className="border border-slate-300 px-3 py-2 text-right text-[13px] font-extrabold text-slate-900 tabular-nums whitespace-nowrap">{won(r.current)}</td>
                        <td className="border border-slate-300 px-3 py-2 text-center whitespace-nowrap">
                          <button
                            onClick={() => setSelectedStoreCode(r.store.code)}
                            className="rounded-lg border border-slate-300 bg-white px-3 py-1 text-sm font-semibold text-black hover:border-blue-300 hover:bg-blue-50"
                          >
                            품목 보기
                          </button>
                        </td>
                      </tr>
                    ))}
                    {!storeRows.length && (
                      <tr>
                        <td colSpan={7} className="border border-slate-300 p-8 text-center text-slate-500">
                          표시할 데이터가 없습니다.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </>
              ) : (
                <>
                  <thead>
                    <tr>
                      <ItemAnalysisSortableTh rowSpan={2} sortKey="brand" sortConfig={storeSortConfig} onSort={requestStoreSort}>브랜드</ItemAnalysisSortableTh>
                      <ItemAnalysisSortableTh rowSpan={2} sortKey="code" sortConfig={storeSortConfig} onSort={requestStoreSort}>거래처코드</ItemAnalysisSortableTh>
                      <ItemAnalysisSortableTh rowSpan={2} sortKey="name" sortConfig={storeSortConfig} onSort={requestStoreSort}>거래처명</ItemAnalysisSortableTh>
                      <ItemAnalysisSortableTh rowSpan={2} sortKey="manager" sortConfig={storeSortConfig} onSort={requestStoreSort}>담당자</ItemAnalysisSortableTh>
                      <PopupTh colSpan={2} className="period-group-start bg-[#F7FCEB] text-center text-black">전년동월</PopupTh>
                      <PopupTh colSpan={2} className="period-group-start bg-[#F3FAFD] text-center text-black">전월</PopupTh>
                      <PopupTh colSpan={1} className="period-group-start bg-[#FFF7FA] text-center text-black">당월</PopupTh>
                      <PopupTh rowSpan={2}>상세</PopupTh>
                    </tr>
                    <tr>
                      <PopupTh top="top-[37px]" className="period-group-start bg-[#F7FCEB] text-right text-black">매출</PopupTh>
                      <ItemAnalysisSortableTh top="top-[37px]" className="bg-[#F7FCEB] text-black" sortKey="prevYearRate" sortConfig={storeSortConfig} onSort={requestStoreSort} right>당월 대비</ItemAnalysisSortableTh>
                      <PopupTh top="top-[37px]" className="period-group-start bg-[#F3FAFD] text-right text-black">매출</PopupTh>
                      <ItemAnalysisSortableTh top="top-[37px]" className="bg-[#F3FAFD] text-black" sortKey="prevMonthRate" sortConfig={storeSortConfig} onSort={requestStoreSort} right>당월 대비</ItemAnalysisSortableTh>
                      <ItemAnalysisSortableTh top="top-[37px]" className="period-group-start bg-[#FFF7FA] text-black" sortKey="current" sortConfig={storeSortConfig} onSort={requestStoreSort} right>매출</ItemAnalysisSortableTh>
                    </tr>
                  </thead>
                  <tbody>
                    {sortedStoreRows.map((r) => (
                      <tr key={r.store.code} className="hover:bg-blue-50">
                        <td className="border border-slate-300 p-2 text-slate-900">{r.store.brand}</td>
                        <td className="border border-slate-300 p-2 text-slate-900">{r.store.code}</td>
                        <td className="border border-slate-300 p-2 font-semibold text-slate-900">{r.store.name}</td>
                        <td className="border border-slate-300 p-2 text-slate-900">{r.store.manager || "미지정"}</td>
                        <td className="border border-slate-300 p-2 text-right font-semibold text-slate-900">{won(r.prevYear)}</td>
                        <td className={`border border-slate-300 p-2 text-right ${itemMetricDiff(r.current, r.prevYear) >= 0 ? "text-emerald-600" : "text-red-600"}`}>{itemSignedPct(itemMetricRate(r.current, r.prevYear))}</td>
                        <td className="border border-slate-300 p-2 text-right font-semibold text-slate-900">{won(r.prevMonth)}</td>
                        <td className={`border border-slate-300 p-2 text-right ${itemMetricDiff(r.current, r.prevMonth) >= 0 ? "text-emerald-600" : "text-red-600"}`}>{itemSignedPct(itemMetricRate(r.current, r.prevMonth))}</td>
                        <td className="border border-slate-300 p-2 text-right font-bold text-slate-900">{won(r.current)}</td>
                        <td className="border border-slate-300 p-2">
                          <button onClick={() => setSelectedStoreCode(r.store.code)} className="rounded-lg border border-slate-300 bg-white px-3 py-1 text-sm font-semibold text-black hover:border-blue-300 hover:bg-blue-50">품목 보기</button>
                        </td>
                      </tr>
                    ))}
                    {!storeRows.length && (
                      <tr>
                        <td colSpan={10} className="border border-slate-300 p-8 text-center text-slate-500">표시할 데이터가 없습니다.</td>
                      </tr>
                    )}
                  </tbody>
                </>
              )}
            </table>
          </div>
        </div>
      )}

      {selectedStoreCode && !selectedItemCode && selectedStore && (
        <div className="overflow-hidden rounded-2xl border border-slate-300 bg-white shadow-sm">
          <div className="flex items-center justify-between border-b border-slate-300 px-4 py-3">
            <div>
              <div className="text-sm font-bold text-slate-800">
                {selectedStore.name} 품목별 비교
              </div>
              <div className="mt-1 text-sm text-slate-500">
                선택한 기간의 당일까지 매출과 전월·전년동월 동일기간 매출을 비교합니다.
              </div>
            </div>
            <div className="flex items-center gap-2">
              <button
                onClick={backToStores}
                className="rounded-lg bg-slate-100 px-3 py-1 text-sm font-semibold text-slate-700 hover:bg-slate-200"
              >
                ← 거래처 목록
              </button>
            </div>
          </div>
          <div className="max-h-[74vh] overflow-auto isolate">
            <table className="w-full min-w-[820px] connected-two-tier border-separate border-spacing-0 text-center text-[13px] whitespace-nowrap">
              <thead>
                <tr>
                  <ItemAnalysisSortableTh
                    rowSpan={2}
                    sortKey="itemCode"
                    sortConfig={itemSortConfig}
                    onSort={requestItemSort}
                  >
                    품목코드
                  </ItemAnalysisSortableTh>
                  <ItemAnalysisSortableTh
                    rowSpan={2}
                    sortKey="itemName"
                    sortConfig={itemSortConfig}
                    onSort={requestItemSort}
                  >
                    품목명
                  </ItemAnalysisSortableTh>
                  <PopupTh
                    colSpan={3}
                    className="period-group-start bg-white text-center font-extrabold text-black"
                  >
                    Time Gone 대비
                  </PopupTh>
                  <PopupTh rowSpan={2}>상세</PopupTh>
                </tr>
                <tr>
                  <ItemAnalysisSortableTh
                    top="top-[37px]"
                    className="period-group-start bg-[#F7FCEB] text-black"
                    sortKey="prevYearSales"
                    sortConfig={itemSortConfig}
                    onSort={requestItemSort}
                    right
                  >
                    전년동월 매출
                  </ItemAnalysisSortableTh>
                  <ItemAnalysisSortableTh
                    top="top-[37px]"
                    className="period-subgroup-start bg-[#F3FAFD] text-black"
                    sortKey="prevMonthSales"
                    sortConfig={itemSortConfig}
                    onSort={requestItemSort}
                    right
                  >
                    전월 매출
                  </ItemAnalysisSortableTh>
                  <ItemAnalysisSortableTh
                    top="top-[37px]"
                    className="period-subgroup-start bg-[#FFF7FA] text-black"
                    sortKey="currentSales"
                    sortConfig={itemSortConfig}
                    onSort={requestItemSort}
                    right
                  >
                    당일까지 매출
                  </ItemAnalysisSortableTh>
                </tr>
              </thead>
              <tbody>
                <tr className="sticky top-[74px] z-10 bg-amber-50 font-extrabold text-slate-900 shadow-sm">
                  <td
                    colSpan={2}
                    className="border border-slate-300 px-3 py-2 text-center"
                  >
                    SUBTOTAL
                  </td>
                  <td className="border border-slate-300 px-3 py-2 text-right">
                    {won(itemSalesSubtotal.prevYearSales)}
                  </td>
                  <td className="border border-slate-300 px-3 py-2 text-right">
                    {won(itemSalesSubtotal.prevMonthSales)}
                  </td>
                  <td className="border border-slate-300 px-3 py-2 text-right text-blue-700">
                    {won(itemSalesSubtotal.currentSales)}
                  </td>
                  <td className="border border-slate-300 px-3 py-2">-</td>
                </tr>
                {sortedItemRows.map((r) => (
                  <tr
                    key={`${r.itemCode}-${r.itemName}`}
                    className="hover:bg-blue-50"
                  >
                    <td className="border border-slate-300 p-2">
                      {r.itemCode}
                    </td>
                    <td className="border border-slate-300 p-2 font-semibold">
                      {r.itemName}
                    </td>
                    <td className="border border-slate-300 p-2 text-right">
                      {won(r.prevYear.sales)}
                    </td>
                    <td className="border border-slate-300 p-2 text-right">
                      {won(r.prevMonth.sales)}
                    </td>
                    <td className="border border-slate-300 p-2 text-right font-bold text-blue-700">
                      {won(r.current.sales)}
                    </td>
                    <td className="border border-slate-300 p-2">
                      <button
                        onClick={() => setSelectedItemCode(r.itemCode)}
                        className="rounded-lg border border-slate-300 bg-white px-3 py-1 text-sm font-semibold text-black transition hover:border-blue-300 hover:bg-blue-50"
                      >
                        상세
                      </button>
                    </td>
                  </tr>
                ))}
                {!sortedItemRows.length && (
                  <tr>
                    <td
                      colSpan={6}
                      className="border border-slate-300 p-8 text-center text-slate-500"
                    >
                      표시할 품목이 없습니다.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {selectedStoreCode &&
        selectedItemCode &&
        selectedItem &&
        selectedStore && (
          <div className="space-y-4">
            <div className="rounded-2xl border border-slate-300 bg-white p-4 shadow-sm">
              <div className="flex items-center justify-between gap-3">
                <div>
                  <div className="text-sm font-bold text-slate-800">
                    {selectedStore.name} / {selectedItem.itemName}
                  </div>
                  <div className="mt-1 text-sm text-slate-500">
                    품목코드 {selectedItem.itemCode} / 현재기간 상세 발주 원본
                  </div>
                </div>
                <button
                  onClick={backToItems}
                  className="rounded-lg bg-slate-100 px-3 py-1 text-sm font-semibold text-slate-700 hover:bg-slate-200"
                >
                  ← 품목 목록
                </button>
              </div>
            </div>
            <div className="grid gap-4 lg:grid-cols-3">
              <ItemCompareCard
                title="수량"
                current={selectedItem.current.qty}
                prev={selectedItem.prevMonth.qty}
                prevYear={selectedItem.prevYear.qty}
              />
              <ItemCompareCard
                title="매출"
                current={selectedItem.current.sales}
                prev={selectedItem.prevMonth.sales}
                prevYear={selectedItem.prevYear.sales}
                money
              />
              <div className="rounded-2xl border border-slate-300 bg-white p-4 text-sm shadow-sm">
                <div className="font-bold text-slate-800">이익</div>
                <div className="mt-3 space-y-2 text-xs">
                  <div className="flex justify-between">
                    <span>현재 이익금액</span>
                    <b>{won(selectedItem.current.profit)}</b>
                  </div>
                  <div className="flex justify-between">
                    <span>현재 이익률</span>
                    <b>{pct(itemMarginRate(selectedItem.current))}</b>
                  </div>
                  <div className="flex justify-between">
                    <span>전월 이익금액</span>
                    <b>{won(selectedItem.prevMonth.profit)}</b>
                  </div>
                  <div className="flex justify-between">
                    <span>전월 이익률</span>
                    <b>{pct(itemMarginRate(selectedItem.prevMonth))}</b>
                  </div>
                  <div className="flex justify-between">
                    <span>전년 이익금액</span>
                    <b>{won(selectedItem.prevYear.profit)}</b>
                  </div>
                  <div className="flex justify-between">
                    <span>전년 이익률</span>
                    <b>{pct(itemMarginRate(selectedItem.prevYear))}</b>
                  </div>
                </div>
              </div>
            </div>
            <div className="overflow-hidden rounded-2xl border border-slate-300 bg-white shadow-sm">
              <div className="border-b border-slate-300 px-4 py-3 text-sm font-bold text-slate-800">
                상세 발주 원본
              </div>
              <div className="max-h-[68vh] overflow-auto isolate">
                <table className="w-full min-w-[1100px] connected-two-tier border-separate border-spacing-0 text-center text-[13px] whitespace-nowrap">
                  <thead>
                    <tr>
                      {(
                        [
                          ["saleDate", "주문일", false],
                          ["storeName", "거래처", false],
                          ["itemCode", "상품코드", false],
                          ["itemName", "상품명", false],
                          ["quantity", "수량", true],
                          ["salesAmount", "매출금액", true],
                          ["costAmount", "원가금액", true],
                          ["profitAmount", "이익금액", true],
                          ["profitRate", "이익률", true],
                        ] as [ItemDetailSortKey, string, boolean][]
                      ).map(([key, label, right]) => (
                        <PopupTh key={key} right={right}>
                          <button
                            type="button"
                            onClick={() => requestDetailSort(key)}
                            className="flex w-full items-center justify-center gap-1 text-center"
                            title="정렬"
                          >
                            <span>{label}</span>
                            <span className="text-[11px] text-black">
                              {sortArrow(
                                detailSortConfig.key === key,
                                detailSortConfig.direction,
                              )}
                            </span>
                          </button>
                        </PopupTh>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {sortedDetailRows.map((r) => (
                      <tr key={r.id} className="hover:bg-blue-50">
                        <td className="border border-gray-300 p-2">
                          {r.saleDate}
                        </td>
                        <td className="border border-gray-300 p-2 font-semibold">
                          {r.storeName}
                        </td>
                        <td className="border border-gray-300 p-2">
                          {r.itemCode}
                        </td>
                        <td className="border border-gray-300 p-2">
                          {r.itemName}
                        </td>
                        <td className="border border-gray-300 p-2 text-right">
                          {won(r.quantity)}
                        </td>
                        <td className="border border-gray-300 p-2 text-right font-bold text-slate-900">
                          {won(r.salesAmount)}
                        </td>
                        <td className="border border-gray-300 p-2 text-right">
                          {won(r.costAmount)}
                        </td>
                        <td className="border border-gray-300 p-2 text-right">
                          {won(r.profitAmount)}
                        </td>
                        <td className="border border-gray-300 p-2 text-right">
                          {pct(r.profitRate)}
                        </td>
                      </tr>
                    ))}
                    {!sortedDetailRows.length && (
                      <tr>
                        <td
                          colSpan={9}
                          className="border border-gray-300 p-8 text-center text-slate-500"
                        >
                          상세 데이터가 없습니다.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}
    </div>
  );
}

function itemCategoryFromName(itemName: string) {
  const name = String(itemName || "").toLowerCase();
  if (/감자|프라이|후렌치|포테이토|웨지|슈스트링|크링클|펍칩/.test(name)) return "감자/프라이";
  if (/번|브리오슈|햄버거빵|빵/.test(name)) return "번/베이커리";
  if (/패티|비프|소고기|치킨패티|새우패티|모짜패티/.test(name)) return "패티";
  if (/치즈|체다|모짜렐라|스위스/.test(name)) return "치즈";
  if (/소스|마요|케첩|머스타드|드레싱/.test(name)) return "소스";
  if (/어니언링|치즈스틱|해쉬브라운|너겟|윙|사이드/.test(name)) return "사이드";
  return "기타";
}

function ItemShipmentAnalysis({
  stores,
  sales,
  itemMasters,
  month,
  date,
}: {
  stores: Store[];
  sales: SalesRecord[];
  itemMasters: ItemMasterRecord[];
  month: string;
  date: string;
}) {
  const [searchDraft, setSearchDraft] = useState("");
  const [search, setSearch] = useState("");
  const [selectedItemCode, setSelectedItemCode] = useState("");
  const [analysisStart, setAnalysisStart] = useState(monthStart(month));
  const [analysisEnd, setAnalysisEnd] = useState(
    date.startsWith(month) ? date : monthEnd(month),
  );
  const [categoryFilter, setCategoryFilter] = useState("전체");
  const [profitRateFilter, setProfitRateFilter] = useState("전체");

  useEffect(() => {
    setAnalysisStart(monthStart(month));
    setAnalysisEnd(date.startsWith(month) ? date : monthEnd(month));
    setSelectedItemCode("");
    setCategoryFilter("전체");
    setProfitRateFilter("전체");
  }, [month, date]);

  const currentStart =
    analysisStart <= analysisEnd ? analysisStart : analysisEnd;
  const currentEnd = analysisStart <= analysisEnd ? analysisEnd : analysisStart;
  const prevStart = sameDayPrevMonth(currentStart);
  const prevEnd = sameDayPrevMonth(currentEnd);
  const normalizedSearch = search.trim().toLowerCase();
  const storeByCode = useMemo(
    () => new Map(stores.map((s) => [s.code, s])),
    [stores],
  );
  const itemMasterByCode = useMemo(
    () => new Map(itemMasters.map((item) => [item.itemCode, item])),
    [itemMasters],
  );

  const itemRows = useMemo(() => {
    const map = new Map<
      string,
      {
        itemCode: string;
        itemName: string;
        category: string;
        current: ItemMetric;
        prevMonth: ItemMetric;
        storeCodes: Set<string>;
      }
    >();

    const ensure = (row: SalesRecord) => {
      const key = row.itemCode || row.itemName || "미지정";
      if (!map.has(key)) {
        map.set(key, {
          itemCode: row.itemCode || "-",
          itemName: row.itemName || "미지정",
          category: itemMasterByCode.get(row.itemCode)?.category || itemCategoryFromName(row.itemName),
          current: emptyItemMetric(),
          prevMonth: emptyItemMetric(),
          storeCodes: new Set<string>(),
        });
      }
      return map.get(key)!;
    };

    sales.forEach((row) => {
      const store = storeByCode.get(row.storeCode);
      const haystack = [
        row.itemCode,
        row.itemName,
        row.storeCode,
        row.storeName,
        store?.name,
        store?.manager,
        store?.channel,
      ]
        .join(" ")
        .toLowerCase();
      if (normalizedSearch && !haystack.includes(normalizedSearch)) return;

      const item = ensure(row);
      if (inRange(row.saleDate, currentStart, currentEnd)) {
        addItemMetric(item.current, row);
        item.storeCodes.add(row.storeCode);
      }
      if (inRange(row.saleDate, prevStart, prevEnd)) {
        addItemMetric(item.prevMonth, row);
        item.storeCodes.add(row.storeCode);
      }
    });

    return Array.from(map.values())
      .filter(
        (r) =>
          r.current.sales ||
          r.current.cost ||
          r.current.profit ||
          r.prevMonth.sales ||
          r.prevMonth.cost ||
          r.prevMonth.profit,
      )
      .map((r) => {
        const prevMonthUnitCost = r.prevMonth.qty
          ? r.prevMonth.cost / r.prevMonth.qty
          : 0;
        const currentUnitCost = r.current.qty
          ? r.current.cost / r.current.qty
          : 0;
        const prevMonthProfitRate = itemMarginRate(r.prevMonth);
        const currentProfitRate = itemMarginRate(r.current);
        return {
          ...r,
          prevMonthUnitCost,
          currentUnitCost,
          prevMonthProfitRate,
          currentProfitRate,
          profitRateChange: currentProfitRate - prevMonthProfitRate,
        };
      })
      .filter((r) => categoryFilter === "전체" || r.category === categoryFilter)
      .filter((r) => {
        if (profitRateFilter === "전체") return true;
        if (profitRateFilter === "30% 미만") return r.currentProfitRate < 30;
        if (profitRateFilter === "30~40%")
          return r.currentProfitRate >= 30 && r.currentProfitRate < 40;
        if (profitRateFilter === "40% 이상") return r.currentProfitRate >= 40;
        if (profitRateFilter === "하락") return r.profitRateChange < 0;
        if (profitRateFilter === "상승") return r.profitRateChange > 0;
        return true;
      })
      .sort((a, b) => b.current.sales - a.current.sales);
  }, [
    sales,
    storeByCode,
    normalizedSearch,
    currentStart,
    currentEnd,
    prevStart,
    prevEnd,
    categoryFilter,
    profitRateFilter,
    itemMasterByCode,
  ]);

  const categoryOptions = useMemo(() => {
    const categories = new Set<string>();
    itemMasters.forEach((item) => {
      if (item.category) categories.add(item.category);
    });
    sales.forEach((row) =>
      categories.add(
        itemMasterByCode.get(row.itemCode)?.category ||
          itemCategoryFromName(row.itemName),
      ),
    );
    return [
      "전체",
      ...Array.from(categories).sort((a, b) =>
        a.localeCompare(b, "ko-KR"),
      ),
    ];
  }, [sales, itemMasters, itemMasterByCode]);

  const subtotal = useMemo(() => {
    const prevMonth = itemRows.reduce(
      (acc, row) => ({
        qty: acc.qty + row.prevMonth.qty,
        sales: acc.sales + row.prevMonth.sales,
        cost: acc.cost + row.prevMonth.cost,
        profit: acc.profit + row.prevMonth.profit,
        profitRateWeighted:
          acc.profitRateWeighted + row.prevMonth.profitRateWeighted,
        profitRateSales: acc.profitRateSales + row.prevMonth.profitRateSales,
      }),
      emptyItemMetric(),
    );
    const current = itemRows.reduce(
      (acc, row) => ({
        qty: acc.qty + row.current.qty,
        sales: acc.sales + row.current.sales,
        cost: acc.cost + row.current.cost,
        profit: acc.profit + row.current.profit,
        profitRateWeighted:
          acc.profitRateWeighted + row.current.profitRateWeighted,
        profitRateSales: acc.profitRateSales + row.current.profitRateSales,
      }),
      emptyItemMetric(),
    );
    const prevRate = itemMarginRate(prevMonth);
    const currentRate = itemMarginRate(current);
    return {
      prevMonth,
      current,
      prevMonthUnitCost: prevMonth.qty ? prevMonth.cost / prevMonth.qty : 0,
      currentUnitCost: current.qty ? current.cost / current.qty : 0,
      prevRate,
      currentRate,
      rateChange: currentRate - prevRate,
      storeCount: new Set(itemRows.flatMap((row) => Array.from(row.storeCodes))).size,
    };
  }, [itemRows]);

  const selectedItem = itemRows.find((r) => r.itemCode === selectedItemCode);

  const storeRows = useMemo(() => {
    if (!selectedItemCode) return [];
    const map = new Map<
      string,
      {
        storeCode: string;
        storeName: string;
        manager: string;
        channel: string;
        current: ItemMetric;
        prevMonth: ItemMetric;
      }
    >();
    const ensure = (row: SalesRecord) => {
      const store = storeByCode.get(row.storeCode);
      const key = row.storeCode || row.storeName || "미지정";
      if (!map.has(key)) {
        map.set(key, {
          storeCode: row.storeCode || "-",
          storeName: store?.name || row.storeName || "미지정",
          manager: store?.manager || row.manager || "미지정",
          channel: store?.channel || row.channel || "미지정",
          current: emptyItemMetric(),
          prevMonth: emptyItemMetric(),
        });
      }
      return map.get(key)!;
    };

    sales
      .filter((row) => (row.itemCode || row.itemName || "미지정") === selectedItemCode)
      .forEach((row) => {
        const storeRow = ensure(row);
        if (inRange(row.saleDate, currentStart, currentEnd))
          addItemMetric(storeRow.current, row);
        if (inRange(row.saleDate, prevStart, prevEnd))
          addItemMetric(storeRow.prevMonth, row);
      });

    return Array.from(map.values())
      .filter(
        (r) =>
          r.current.sales ||
          r.current.cost ||
          r.current.profit ||
          r.prevMonth.sales ||
          r.prevMonth.cost ||
          r.prevMonth.profit,
      )
      .sort((a, b) => b.current.sales - a.current.sales);
  }, [sales, selectedItemCode, storeByCode, currentStart, currentEnd, prevStart, prevEnd]);

  function applySearch() {
    setSearch(searchDraft.trim());
    setSelectedItemCode("");
  }

  return (
    <div className="flex min-h-0 flex-1 flex-col space-y-3">
      <div className="bg-white py-1">
        <div className="flex flex-wrap items-end justify-end gap-3">
          <div>
            <div className="text-base font-bold text-black">품목분석</div>
            <div className="mt-1 text-xs text-black">
              당월 {currentStart} ~ {currentEnd} / 전월 {prevStart} ~ {prevEnd}
            </div>
          </div>
          <div className="flex flex-wrap items-center gap-2 text-xs text-black">
            <input
              value={searchDraft}
              onChange={(e) => {
                const next = e.target.value;
                setSearchDraft(next);
                if (!next.trim()) setSearch("");
              }}
              onKeyDown={(e) => {
                if (e.key === "Enter") applySearch();
              }}
              placeholder="품목/거래처/담당자 검색 후 Enter"
              className="h-8 w-[260px] rounded-lg border border-slate-300 bg-white px-3 text-xs text-black outline-none focus:border-blue-500"
            />
            <button
              type="button"
              onClick={applySearch}
              className="h-8 rounded-lg bg-slate-800 px-3 text-xs font-semibold text-white hover:bg-slate-700"
            >
              검색
            </button>
            <label className="flex items-center gap-1 text-black">
              시작일
              <input
                type="date"
                value={analysisStart}
                onChange={(e) => setAnalysisStart(e.target.value)}
                className="h-8 rounded-lg border border-slate-300 px-2 text-xs text-black outline-none focus:border-blue-500"
              />
            </label>
            <label className="flex items-center gap-1 text-black">
              종료일
              <input
                type="date"
                value={analysisEnd}
                onChange={(e) => setAnalysisEnd(e.target.value)}
                className="h-8 rounded-lg border border-slate-300 px-2 text-xs text-black outline-none focus:border-blue-500"
              />
            </label>
          </div>
        </div>
      </div>

      {!selectedItemCode && (
        <div className="overflow-hidden rounded-2xl border border-slate-300 bg-white shadow-sm">
          <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-300 px-4 py-3">
            <div className="text-sm font-bold text-black">품목별 손익 요약</div>
            <div className="text-xs font-semibold text-black">
              필터 결과 {itemRows.length.toLocaleString("ko-KR")}개 품목
            </div>
          </div>
          <div className="overflow-x-auto isolate">
            <div className="min-w-[1650px]">
              {/* 헤더와 SUBTOTAL은 스크롤 영역 밖에 두어 완전히 고정합니다. */}
              <div className="item-profit-fixed-header-wrap">
              <table className="item-profit-fixed-header w-full table-fixed text-center text-black whitespace-nowrap">
              <colgroup>
                <col style={{ width: "120px" }} />
                <col style={{ width: "300px" }} />
                <col style={{ width: "100px" }} />
                <col style={{ width: "130px" }} />
                <col style={{ width: "100px" }} />
                <col style={{ width: "120px" }} />
                <col style={{ width: "82px" }} />
                <col style={{ width: "130px" }} />
                <col style={{ width: "100px" }} />
                <col style={{ width: "120px" }} />
                <col style={{ width: "90px" }} />
                <col style={{ width: "82px" }} />
                <col style={{ width: "78px" }} />
                <col style={{ width: "100px" }} />
              </colgroup>
              <thead>
                <tr className="bg-slate-100">
                  <th rowSpan={2} className="border border-slate-300 bg-white px-2 py-2 font-bold text-black">품목코드</th>
                  <th rowSpan={2} className="border border-slate-300 bg-white px-3 py-2 font-bold text-black">품목명</th>
                  <th rowSpan={2} className="border border-slate-300 bg-white px-1 py-2 font-bold text-black">
                    <div className="flex w-full flex-col items-center gap-1">
                      <span>카테고리</span>
                      <select
                        value={categoryFilter}
                        onChange={(e) => {
                          setCategoryFilter(e.target.value);
                          setSelectedItemCode("");
                        }}
                        className="w-full rounded border border-slate-300 bg-white px-1 font-semibold text-black"
                      >
                        {categoryOptions.map((category) => (
                          <option key={category} value={category}>{category}</option>
                        ))}
                      </select>
                    </div>
                  </th>
                  <th colSpan={4} className="border border-slate-300 bg-[#F3FAFD] px-3 py-1 text-[15px] font-extrabold text-black">전월</th>
                  <th colSpan={4} className="border border-slate-300 bg-[#FFF7FA] px-3 py-1 text-[15px] font-extrabold text-black">당월</th>
                  <th rowSpan={2} className="border border-slate-300 bg-[#FFF9F3] px-1 py-2 font-bold text-black">이익률변동</th>
                  <th rowSpan={2} className="border border-slate-300 bg-white px-1 py-2 font-bold text-black">사용 거래처 수</th>
                  <th rowSpan={2} className="border border-slate-300 bg-white px-2 py-2 font-bold text-black">상세</th>
                </tr>
                <tr>
                  <th className="border border-slate-300 bg-[#F3FAFD] px-2 py-2 font-bold text-black">매출</th>
                  <th className="border border-slate-300 bg-[#F3FAFD] px-2 py-2 font-bold text-black">매입단가</th>
                  <th className="border border-slate-300 bg-[#F3FAFD] px-2 py-2 text-[14px] font-bold text-black">이익금액</th>
                  <th className="border border-slate-300 bg-[#F3FAFD] px-1 py-2 font-bold text-black">이익률</th>
                  <th className="border border-slate-300 bg-[#FFF7FA] px-2 py-2 font-bold text-black">매출</th>
                  <th className="border border-slate-300 bg-[#FFF7FA] px-2 py-2 font-bold text-black">매입단가</th>
                  <th className="border border-slate-300 bg-[#FFF7FA] px-2 py-2 text-[14px] font-bold text-black">이익금액</th>
                  <th className="border border-slate-300 bg-[#FFF7FA] px-1 py-1 font-bold text-black">
                    <div className="flex w-full flex-col items-center gap-1">
                      <span>이익률</span>
                      <select
                        value={profitRateFilter}
                        onChange={(e) => {
                          setProfitRateFilter(e.target.value);
                          setSelectedItemCode("");
                        }}
                        className="w-full rounded border border-slate-300 bg-white px-1 font-semibold text-black"
                      >
                        {["전체", "30% 미만", "30~40%", "40% 이상", "상승", "하락"].map((value) => (
                          <option key={value} value={value}>{value}</option>
                        ))}
                      </select>
                    </div>
                  </th>
                </tr>
                <tr className="item-profit-subtotal font-extrabold text-black">
                  <th colSpan={3} className="subtotal-label border border-slate-400 font-extrabold">SUBTOTAL</th>
                  <th className="subtotal-number sales-value-cell border border-slate-400 font-extrabold">{won(subtotal.prevMonth.sales)}</th>
                  <th className="subtotal-number border border-slate-400">{won(subtotal.prevMonthUnitCost)}</th>
                  <th className="subtotal-number border border-slate-400">{won(subtotal.prevMonth.profit)}</th>
                  <th className="subtotal-number border border-slate-400">{pct(subtotal.prevRate)}</th>
                  <th className="subtotal-number sales-value-cell border border-slate-400 font-extrabold">{won(subtotal.current.sales)}</th>
                  <th className="subtotal-number border border-slate-400">{won(subtotal.currentUnitCost)}</th>
                  <th className="subtotal-number border border-slate-400">{won(subtotal.current.profit)}</th>
                  <th className="subtotal-number border border-slate-400">{pct(subtotal.currentRate)}</th>
                  <th className={`subtotal-number border border-slate-400 ${subtotal.rateChange > 0 ? "text-emerald-700" : subtotal.rateChange < 0 ? "text-red-600" : "text-black"}`}>
                    {itemSignedPct(subtotal.rateChange)}
                  </th>
                  <th className="border border-slate-400 text-center">{subtotal.storeCount.toLocaleString("ko-KR")}</th>
                  <th className="border border-slate-400" />
                </tr>
              </thead>
              </table>
              </div>

              {/* 품목 데이터 행만 세로로 스크롤됩니다. */}
              <div className="max-h-[68vh] overflow-y-scroll overflow-x-hidden">
                <table className="item-profit-fixed-body w-full table-fixed text-center text-black whitespace-nowrap">
              <colgroup>
                <col style={{ width: "120px" }} />
                <col style={{ width: "300px" }} />
                <col style={{ width: "100px" }} />
                <col style={{ width: "130px" }} />
                <col style={{ width: "100px" }} />
                <col style={{ width: "120px" }} />
                <col style={{ width: "82px" }} />
                <col style={{ width: "130px" }} />
                <col style={{ width: "100px" }} />
                <col style={{ width: "120px" }} />
                <col style={{ width: "90px" }} />
                <col style={{ width: "82px" }} />
                <col style={{ width: "78px" }} />
                <col style={{ width: "100px" }} />
              </colgroup>
              <tbody>
                {itemRows.map((r) => (
                  <tr key={`${r.itemCode}-${r.itemName}`} className="hover:bg-blue-50">
                    <td className="border border-slate-300 p-2">{r.itemCode}</td>
                    <td className="item-name-cell border border-slate-300 p-2 text-left font-semibold" title={r.itemName}>{r.itemName}</td>
                    <td className="border border-slate-300 p-2 font-semibold">{r.category}</td>
                    <td className="item-profit-number-cell sales-value-cell border border-slate-300 p-2 font-bold">{won(r.prevMonth.sales)}</td>
                    <td className="item-profit-number-cell border border-slate-300 p-2">{won(r.prevMonthUnitCost)}</td>
                    <td className="item-profit-number-cell border border-slate-300 p-2">{won(r.prevMonth.profit)}</td>
                    <td className="item-profit-number-cell border border-slate-300 p-2 font-bold">{pct(r.prevMonthProfitRate)}</td>
                    <td className="item-profit-number-cell sales-value-cell border border-slate-300 p-2 font-extrabold">{won(r.current.sales)}</td>
                    <td className="item-profit-number-cell border border-slate-300 p-2">{won(r.currentUnitCost)}</td>
                    <td className="item-profit-number-cell border border-slate-300 p-2 font-bold">{won(r.current.profit)}</td>
                    <td className="item-profit-number-cell border border-slate-300 p-2 font-extrabold">{pct(r.currentProfitRate)}</td>
                    <td className={`item-profit-number-cell border border-slate-300 p-2 font-extrabold ${r.profitRateChange > 0 ? "text-emerald-700" : r.profitRateChange < 0 ? "text-red-600" : "text-black"}`}>
                      {itemSignedPct(r.profitRateChange)}
                    </td>
                    <td className="border border-slate-300 p-2 text-center font-extrabold">{r.storeCodes.size.toLocaleString("ko-KR")}</td>
                    <td className="border border-slate-300 p-2">
                      <button
                        onClick={() => setSelectedItemCode(r.itemCode)}
                        className="rounded-lg border border-slate-300 bg-white px-3 py-1 text-xs font-semibold text-black hover:border-blue-300 hover:bg-blue-50"
                      >
                        거래처 보기
                      </button>
                    </td>
                  </tr>
                ))}
                {!itemRows.length && (
                  <tr>
                    <td colSpan={14} className="border border-slate-300 p-8 text-center text-black">
                      표시할 품목이 없습니다.
                    </td>
                  </tr>
                )}
              </tbody>
                </table>
              </div>
            </div>
          </div>
        </div>
      )}

      {selectedItemCode && selectedItem && (
        <div className="overflow-hidden rounded-2xl border border-slate-300 bg-white shadow-sm">
          <div className="flex items-center justify-between border-b border-slate-300 px-4 py-3">
            <div>
              <div className="text-sm font-bold text-black">
                {selectedItem.itemName} 거래처별 손익 현황
              </div>
              <div className="mt-1 text-xs text-black">
                품목코드 {selectedItem.itemCode} / 선택 기간 기준 거래처별 매출·이익입니다.
              </div>
            </div>
            <button
              onClick={() => setSelectedItemCode("")}
              className="rounded-lg bg-slate-100 px-3 py-1 text-xs font-semibold text-black hover:bg-slate-200"
            >
              ← 품목 목록
            </button>
          </div>
          <div className="max-h-[68vh] overflow-auto isolate">
            <table className="w-full min-w-[1450px] border-separate border-spacing-0 text-center text-[12px] text-black whitespace-nowrap">
              <thead>
                <tr>
                  <PopupTh>거래처코드</PopupTh>
                  <PopupTh>거래처명</PopupTh>
                  <PopupTh>담당자</PopupTh>
                  <PopupTh>채널</PopupTh>
                  <PopupTh right>전월매출</PopupTh>
                  <PopupTh right>전월이익금액</PopupTh>
                  <PopupTh right>전월이익률</PopupTh>
                  <PopupTh right>당월매출</PopupTh>
                  <PopupTh right>당월이익금액</PopupTh>
                  <PopupTh right>당월이익률</PopupTh>
                  <PopupTh right>이익률변동</PopupTh>
                </tr>
              </thead>
              <tbody>
                {storeRows.map((r) => {
                  const prevRate = itemMarginRate(r.prevMonth);
                  const currentRate = itemMarginRate(r.current);
                  const change = currentRate - prevRate;
                  return (
                    <tr key={r.storeCode} className="hover:bg-blue-50">
                      <td className="border border-slate-300 p-2">{r.storeCode}</td>
                      <td className="border border-slate-300 p-2 text-left font-semibold">{r.storeName}</td>
                      <td className="border border-slate-300 p-2">{r.manager}</td>
                      <td className="border border-slate-300 p-2">{r.channel}</td>
                      <td className="border border-slate-300 p-2 text-right text-[14px] font-bold">{won(r.prevMonth.sales)}</td>
                      <td className="border border-slate-300 p-2 text-right">{won(r.prevMonth.profit)}</td>
                      <td className="border border-slate-300 p-2 text-right">{pct(prevRate)}</td>
                      <td className="border border-slate-300 p-2 text-right text-[14px] font-extrabold">{won(r.current.sales)}</td>
                      <td className="border border-slate-300 p-2 text-right font-bold">{won(r.current.profit)}</td>
                      <td className="border border-slate-300 p-2 text-right font-extrabold">{pct(currentRate)}</td>
                      <td className={`border border-slate-300 p-2 text-right font-extrabold ${change > 0 ? "text-emerald-700" : change < 0 ? "text-red-600" : "text-black"}`}>
                        {itemSignedPct(change)}
                      </td>
                    </tr>
                  );
                })}
                {!storeRows.length && (
                  <tr>
                    <td colSpan={11} className="border border-slate-300 p-8 text-center text-black">
                      표시할 거래처가 없습니다.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

function ItemCompareCard({
  title,
  current,
  prev,
  prevYear,
  money,
}: {
  title: string;
  current: number;
  prev: number;
  prevYear: number;
  money?: boolean;
}) {
  return (
    <div className="rounded-2xl border border-slate-300 bg-white p-4 text-sm shadow-sm">
      <div className="font-bold text-slate-800">{title}</div>
      <div className="mt-3 space-y-2 text-xs">
        <div className="flex justify-between">
          <span>현재</span>
          <b className="text-blue-700">{money ? won(current) : won(current)}</b>
        </div>
        <div className="flex justify-between">
          <span>전월</span>
          <b>{money ? won(prev) : won(prev)}</b>
        </div>
        <div className="flex justify-between">
          <span>전월 차이</span>
          <b
            className={
              itemMetricDiff(current, prev) >= 0
                ? "text-emerald-600"
                : "text-red-600"
            }
          >
            {itemSignedNumber(itemMetricDiff(current, prev), money)}
          </b>
        </div>
        <div className="flex justify-between">
          <span>전월 증감</span>
          <b
            className={
              itemMetricDiff(current, prev) >= 0
                ? "text-emerald-600"
                : "text-red-600"
            }
          >
            {itemSignedPct(itemMetricRate(current, prev))}
          </b>
        </div>
        <div className="flex justify-between">
          <span>전년</span>
          <b>{money ? won(prevYear) : won(prevYear)}</b>
        </div>
        <div className="flex justify-between">
          <span>전년 차이</span>
          <b
            className={
              itemMetricDiff(current, prevYear) >= 0
                ? "text-emerald-600"
                : "text-red-600"
            }
          >
            {itemSignedNumber(itemMetricDiff(current, prevYear), money)}
          </b>
        </div>
        <div className="flex justify-between">
          <span>전년 증감</span>
          <b
            className={
              itemMetricDiff(current, prevYear) >= 0
                ? "text-emerald-600"
                : "text-red-600"
            }
          >
            {itemSignedPct(itemMetricRate(current, prevYear))}
          </b>
        </div>
      </div>
    </div>
  );
}

function SalesStatus({
  stores,
  sales,
  targets,
  ests,
  month,
  date,
  timeGone,
  codeMappings,
  compact = false,
  defaultView = "거래처별",
}: {
  stores: Store[];
  sales: SalesRecord[];
  targets: TargetRecord[];
  ests: EstRecord[];
  month: string;
  date: string;
  timeGone: ReturnType<typeof getTimeGone>;
  codeMappings: StoreCodeMapping[];
  compact?: boolean;
  defaultView?: SalesView;
}) {
  const [view, setView] = useState<SalesView>(defaultView);
  const [search, setSearch] = useState("");
  const [searchDraft, setSearchDraft] = useState("");
  const [drill, setDrill] = useState<{
    title: string;
    rows: SalesRecord[];
  } | null>(null);
  const [sortConfig, setSortConfig] = useState<{
    key: SalesStatusSortKey;
    direction: SortDirection;
  }>({ key: "currentSales", direction: "desc" });
  const [inactiveOpen, setInactiveOpen] = useState(false);
  const [orderDateFilter, setOrderDateFilter] = useState<
    "all" | "check7" | "no30"
  >("all");
  const [hideEndedStores, setHideEndedStores] = useState(false);
  const [selectedManagers, setSelectedManagers] = useState<string[]>([]);

  const normalizedSearch = search.trim().toLowerCase();
  const stMap = storeMap(stores);

  const toggleSelectedManager = (manager: string) => {
    setSelectedManagers((prev) =>
      prev.includes(manager)
        ? prev.filter((item) => item !== manager)
        : [...prev, manager],
    );
  };

  const currentCanonicalStores = useMemo(() => {
    const map = new Map<string, ReturnType<typeof resolveStoreInfo>>();
    sales
      .filter(
        (r) =>
          r.period === "current" &&
          inRange(r.saleDate, monthStart(month), monthEnd(month)),
      )
      .forEach((r) => {
        if (!map.has(r.storeCode)) {
          map.set(
            r.storeCode,
            resolveStoreInfo(
              r.storeCode,
              r.storeName,
              {
                channel: r.channel,
                manager: r.manager,
                storeType: r.storeType,
                brand: r.brand,
              },
              stores,
            ),
          );
        }
      });
    return Array.from(map.values());
  }, [sales, month, stores]);

  const currentByCode = useMemo(
    () => new Map(currentCanonicalStores.map((s) => [norm(s.code), s])),
    [currentCanonicalStores],
  );
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
      brand: displayBrand(mappedStore?.brand || r.brand),
      status: mappedStore?.status || ("거래중" as const),
      originalCode: r.storeCode,
      originalName: r.storeName,
    };
  };

  const resolveRecord = (r: SalesRecord) => {
    if (r.period === "prevYear" || r.period === "prevMonth") {
      const manual = findManualMapping(r.storeCode, r.storeName);
      if (manual?.currentCode) {
        const mapped = resolveStoreInfo(
          manual.currentCode,
          manual.currentName || r.storeName,
          {
            channel: r.channel,
            manager: r.manager,
            storeType: r.storeType,
            brand: r.brand,
          },
          stores,
        );
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

    return resolveStoreInfo(
      r.storeCode,
      r.storeName,
      {
        channel: r.channel,
        manager: r.manager,
        storeType: r.storeType,
        brand: r.brand,
      },
      stores,
    );
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
      String(s.manager || "")
        .toLowerCase()
        .includes(normalizedSearch) ||
      s.channel.toLowerCase().includes(normalizedSearch) ||
      resolved.name.toLowerCase().includes(normalizedSearch) ||
      resolved.code.toLowerCase().includes(normalizedSearch) ||
      resolved.brand.toLowerCase().includes(normalizedSearch) ||
      String(resolved.manager || "")
        .toLowerCase()
        .includes(normalizedSearch) ||
      resolved.channel.toLowerCase().includes(normalizedSearch)
    );
  };

  const shouldIncludeRecord = (s: SalesRecord) =>
    !hideEndedStores || resolveRecord(s).status !== "거래종료";

  const current = sales.filter(
    (s) =>
      s.period === "current" &&
      inRange(s.saleDate, monthStart(month), date) &&
      filterByStoreSearch(s) &&
      shouldIncludeRecord(s),
  );
  const currentFullMonthRows = sales.filter(
    (s) =>
      s.period === "current" &&
      inRange(s.saleDate, monthStart(month), monthEnd(month)) &&
      filterByStoreSearch(s) &&
      shouldIncludeRecord(s),
  );
  const prevMonthRows = sales.filter(
    (s) =>
      s.period === "prevMonth" &&
      s.refMonth === month &&
      filterByStoreSearch(s) &&
      shouldIncludeRecord(s),
  );
  const prevYearRows = sales.filter(
    (s) =>
      s.period === "prevYear" &&
      s.refMonth === month &&
      filterByStoreSearch(s) &&
      shouldIncludeRecord(s),
  );

  const isStoreListView = view === "거래처별" || view === "담당자별";
  const showChannelColumn = view === "거래처별" || view === "브랜드별";

  const rowKey = (r: SalesRecord) => {
    const resolved = resolveRecord(r);
    if (isStoreListView) return resolved.code || resolved.name;
    if (view === "브랜드별") return displayBrand(resolved.brand);
    return resolved.channel || "미지정";
  };

  const rowLabel = (key: string, records: SalesRecord[]) => {
    if (!isStoreListView) return key || "미지정";
    const first = records[0];
    if (first) return resolveRecord(first).name;
    const mapped = stMap.get(key);
    return mapped?.name || key || "미지정";
  };

  const storeKey = (store: Store) => {
    const resolved = resolveStoreInfo(store.code, store.name, store, stores);
    if (isStoreListView) return resolved.code || resolved.name;
    if (view === "브랜드별") return displayBrand(resolved.brand);
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
      const resolved = resolveStoreInfo(
        e.storeCode,
        e.storeName,
        mappedStore || {},
        stores,
      );
      const display = resolved;
      const mappedStoreName = `${mappedStore?.name || ""} ${display.name}`;
      const mappedBrand = `${mappedStore?.brand || ""} ${display.brand}`;
      const mappedManager = `${mappedStore?.manager || ""} ${display.manager}`;
      const mappedChannel = `${mappedStore?.channel || ""} ${display.channel}`;
      const isEnded =
        display.status === "거래종료" || mappedStore?.status === "거래종료";
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
      const resolved = resolveStoreInfo(
        e.storeCode,
        e.storeName,
        s || {},
        stores,
      );
      const display = resolved;
      const key = s
        ? storeKey(s)
        : isStoreListView
          ? display.code || display.name || "미지정"
          : view === "브랜드별"
            ? displayBrand(display.brand)
            : display.channel || "미지정";
      estMap.set(key, (estMap.get(key) || 0) + e.amount);
    });

  const baseKeySet = new Set([
    ...currentMap.keys(),
    ...currentFullMonthMap.keys(),
    ...prevMonthMap.keys(),
    ...prevYearMap.keys(),
    ...estMap.keys(),
  ]);
  if (isStoreListView && !normalizedSearch) {
    const selectedManagerSet = new Set(selectedManagers);
    stores
      .filter((store) => !hideEndedStores || store.status === "거래중")
      .filter(
        (store) =>
          view !== "담당자별" ||
          selectedManagers.length === 0 ||
          selectedManagerSet.has(store.manager || "미지정"),
      )
      .forEach((store) => baseKeySet.add(storeKey(store)));
  }
  const keys = Array.from(baseKeySet).sort();

  const lastOrderDateByKey = useMemo(() => {
    const map = new Map<string, string>();
    sales
      .filter(
        (r) =>
          r.period === "current" &&
          r.saleDate <= date &&
          shouldIncludeRecord(r),
      )
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

  const rows = keys
    .map((key) => {
      const currentRecords = currentMap.get(key) || [];
      const fullMonthRecords = currentFullMonthMap.get(key) || [];
      const prevMonthRecords = prevMonthMap.get(key) || [];
      const prevYearRecords = prevYearMap.get(key) || [];
      const allRecords = [
        ...currentRecords,
        ...fullMonthRecords,
        ...prevMonthRecords,
        ...prevYearRecords,
      ];
      const currentSales = sum(currentRecords, "salesAmount");
      const fullMonthSales = sum(fullMonthRecords, "salesAmount");
      const prevMonthSales = sum(prevMonthRecords, "salesAmount");
      const prevYearSales = sum(prevYearRecords, "salesAmount");
      const profitAmount = sum(currentRecords, "profitAmount");
      const profitRate = weightedProfitRate(currentRecords);
      const prevMonthRate = prevMonthSales
        ? ((currentSales - prevMonthSales) / prevMonthSales) * 100
        : 0;
      const prevYearRate = prevYearSales
        ? ((currentSales - prevYearSales) / prevYearSales) * 100
        : 0;
      const prevYearTimeGoneGap = prevYearRate - timeGone.timeGoneRate;
      const prevMonthTimeGoneGap = prevMonthRate - timeGone.timeGoneRate;
      const est = estMap.get(key) || 0;
      const estRate = est ? (currentSales / est) * 100 : 0;
      const firstRecord = allRecords[0];
      const storeStatus = firstRecord
        ? resolveRecord(firstRecord).status
        : stMap.get(key)?.status;
      const isEndedStore = isStoreListView && storeStatus === "거래종료";
      const resolvedStores = allRecords.map(resolveRecord);
      const uniqueStoreCount = new Set(resolvedStores.map((store) => store.code || store.name).filter(Boolean)).size;
      const channelValues = Array.from(new Set(resolvedStores.map((store) => store.storeType === "매장" ? "매장" : "비매장")));
      const channel = channelValues.length === 1 ? channelValues[0] : channelValues.length > 1 ? "혼합" : (firstRecord ? (resolveRecord(firstRecord).storeType === "매장" ? "매장" : "비매장") : "-");
      const baseLabel = rowLabel(key, allRecords);
      return {
        key,
        label: view === "브랜드별" ? `${baseLabel} (${uniqueStoreCount})` : baseLabel,
        channel,
        prevYearSales,
        prevMonthSales,
        currentSales,
        fullMonthSales,
        prevMonthRate,
        prevYearRate,
        prevYearTimeGoneGap,
        prevMonthTimeGoneGap,
        timeGone: timeGone.timeGoneRate,
        timeGoneGap: estRate - timeGone.timeGoneRate,
        est,
        estRate,
        profitAmount,
        profitRate,
        isEndedStore,
        manager: firstRecord
          ? resolveRecord(firstRecord).manager || "미지정"
          : stMap.get(key)?.manager || "미지정",
        lastOrderDate: isStoreListView
          ? lastOrderDateByKey.get(key) || "-"
          : "-",
        daysSinceLastOrder: isStoreListView
          ? daysBetween(lastOrderDateByKey.get(key) || "-", date)
          : 0,
      };
    })
    .filter((row) => {
      if (!isStoreListView) return true;
      if (orderDateFilter === "check7") return row.daysSinceLastOrder >= 7;
      if (orderDateFilter === "no30") return row.daysSinceLastOrder >= 30;
      return true;
    });

  const displayRows = useMemo(() => {
    let nextRows = rows;
    if (hideEndedStores && isStoreListView)
      nextRows = nextRows.filter((row) => !row.isEndedStore);
    if (view === "담당자별" && selectedManagers.length > 0) {
      const selected = new Set(selectedManagers);
      nextRows = nextRows.filter((row) =>
        selected.has(row.manager || "미지정"),
      );
    }
    return nextRows;
  }, [rows, hideEndedStores, view, selectedManagers]);

  const managerFilterOptions = useMemo(() => {
    if (view !== "담당자별") return [];
    const options = new Set<string>();
    stores.forEach((store) => {
      if (!hideEndedStores || store.status === "거래중")
        options.add(store.manager || "미지정");
    });
    sales.forEach((row) => {
      const resolved = resolveRecord(row);
      if (!hideEndedStores || resolved.status !== "거래종료")
        options.add(resolved.manager || "미지정");
    });
    return Array.from(options)
      .filter(Boolean)
      .sort((a, b) => a.localeCompare(b, "ko-KR"));
  }, [sales, stores, view, hideEndedStores, codeMappings]);

  const managerStoreRowsByKey = useMemo(() => {
    const managerMap = new Map<
      string,
      Map<
        string,
        {
          code: string;
          name: string;
          brand: string;
          channel: string;
          storeType: string;
          prevYearSales: number;
          prevMonthSales: number;
          currentSales: number;
          fullMonthSales: number;
          profitAmount: number;
        }
      >
    >();

    if (view !== "담당자별")
      return new Map<
        string,
        {
          code: string;
          name: string;
          brand: string;
          channel: string;
          storeType: string;
          prevYearSales: number;
          prevMonthSales: number;
          currentSales: number;
          fullMonthSales: number;
          profitAmount: number;
        }[]
      >();

    const ensure = (
      record: SalesRecord,
      amountKey:
        | "prevYearSales"
        | "prevMonthSales"
        | "currentSales"
        | "fullMonthSales",
      includeProfit = false,
    ) => {
      const resolved = resolveRecord(record);
      const managerKey = resolved.manager || "미지정";
      const storeKeyValue = resolved.code || resolved.name || "미지정";
      if (!managerMap.has(managerKey)) managerMap.set(managerKey, new Map());
      const storeMapForManager = managerMap.get(managerKey)!;
      const prev = storeMapForManager.get(storeKeyValue) || {
        code: resolved.code || record.storeCode || "미지정",
        name: resolved.name || record.storeName || "미지정",
        brand: displayBrand(resolved.brand),
        channel: resolved.channel || "미지정",
        storeType: resolved.storeType || "미지정",
        prevYearSales: 0,
        prevMonthSales: 0,
        currentSales: 0,
        fullMonthSales: 0,
        profitAmount: 0,
      };
      prev[amountKey] += Number(record.salesAmount || 0);
      if (includeProfit) prev.profitAmount += Number(record.profitAmount || 0);
      storeMapForManager.set(storeKeyValue, prev);
    };

    prevYearRows.forEach((record) => ensure(record, "prevYearSales"));
    prevMonthRows.forEach((record) => ensure(record, "prevMonthSales"));
    current.forEach((record) => ensure(record, "currentSales", true));
    currentFullMonthRows.forEach((record) => ensure(record, "fullMonthSales"));

    const result = new Map<
      string,
      {
        code: string;
        name: string;
        brand: string;
        channel: string;
        storeType: string;
        prevYearSales: number;
        prevMonthSales: number;
        currentSales: number;
        fullMonthSales: number;
        profitAmount: number;
      }[]
    >();

    managerMap.forEach((value, key) => {
      result.set(
        key,
        Array.from(value.values()).sort(
          (a, b) =>
            b.currentSales - a.currentSales ||
            a.name.localeCompare(b.name, "ko-KR"),
        ),
      );
    });

    return result;
  }, [
    view,
    prevYearRows,
    prevMonthRows,
    current,
    currentFullMonthRows,
    stores,
    codeMappings,
  ]);

  const sortedRows = useMemo(() => {
    return [...displayRows].sort((a, b) => {
      const aValue = a[sortConfig.key];
      const bValue = b[sortConfig.key];
      let result = 0;
      if (sortConfig.key === "lastOrderDate") {
        const aDate = String(aValue || "-");
        const bDate = String(bValue || "-");
        if (aDate === "-" && bDate === "-") result = 0;
        else if (aDate === "-") return 1;
        else if (bDate === "-") return -1;
        else result = aDate.localeCompare(bDate);
      } else if (typeof aValue === "string" || typeof bValue === "string") {
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

  const openDrill = (
    row: { key: string; label: string },
    period: DrillPeriod,
  ) => {
    const rows = getDrillRows(row.key, period);
    setDrill({ title: `${row.label} · ${drillPeriodLabel(period)}`, rows });
  };

  const openTotalDrill = (period: DrillPeriod) => {
    const rows = getTotalDrillRows(period);
    setDrill({ title: `${view} 전체 · ${drillPeriodLabel(period)}`, rows });
  };

  const filteredCurrentSales = displayRows.reduce(
    (a, b) => a + b.currentSales,
    0,
  );
  const filteredFullMonthSales = displayRows.reduce(
    (a, b) => a + b.fullMonthSales,
    0,
  );
  const filteredPrevMonthSales = displayRows.reduce(
    (a, b) => a + b.prevMonthSales,
    0,
  );
  const filteredPrevYearSales = displayRows.reduce(
    (a, b) => a + b.prevYearSales,
    0,
  );
  const filteredProfitAmount = displayRows.reduce(
    (a, b) => a + b.profitAmount,
    0,
  );

  const buildManagerLastOrderRows = (managers: string[]) => {
    const selected = new Set(managers);
    const currentMonthRows = sales.filter(
      (row) =>
        row.period === "current" &&
        inRange(row.saleDate, monthStart(month), date),
    );
    const currentFullRows = sales.filter(
      (row) =>
        row.period === "current" &&
        inRange(row.saleDate, monthStart(month), monthEnd(month)),
    );
    const currentByStore = groupByKey(
      currentMonthRows,
      (row) => resolveRecord(row).code || resolveRecord(row).name || "미지정",
    );
    const fullByStore = groupByKey(
      currentFullRows,
      (row) => resolveRecord(row).code || resolveRecord(row).name || "미지정",
    );
    const latestByStore = new Map<string, string>();

    sales
      .filter((row) => row.period === "current" && row.saleDate <= date)
      .forEach((row) => {
        const resolved = resolveRecord(row);
        const key = resolved.code || resolved.name || "미지정";
        const prev = latestByStore.get(key);
        if (!prev || row.saleDate > prev) latestByStore.set(key, row.saleDate);
      });

    const storeCandidates = new Map<
      string,
      ReturnType<typeof resolveStoreInfo>
    >();
    stores.forEach((store) => {
      const resolved = resolveStoreInfo(store.code, store.name, store, stores);
      if (selected.has(resolved.manager || "미지정"))
        storeCandidates.set(
          resolved.code || resolved.name || "미지정",
          resolved,
        );
    });
    sales.forEach((row) => {
      const resolved = resolveRecord(row);
      if (selected.has(resolved.manager || "미지정"))
        storeCandidates.set(
          resolved.code || resolved.name || "미지정",
          resolved,
        );
    });

    return Array.from(storeCandidates.values())
      .filter((store) => !hideEndedStores || store.status !== "거래종료")
      .map((store) => {
        const key = store.code || store.name || "미지정";
        const lastOrderDate = latestByStore.get(key) || "-";
        const currentRowsForStore = currentByStore.get(key) || [];
        const fullRowsForStore = fullByStore.get(key) || [];
        return {
          담당자: store.manager || "미지정",
          거래처코드: store.code,
          거래처명: store.name,
          마지막발주일: lastOrderDate,
          경과일:
            lastOrderDate === "-" ? "-" : daysBetween(lastOrderDate, date),
          당월매출: sum(currentRowsForStore, "salesAmount"),
          당월전체매출: sum(fullRowsForStore, "salesAmount"),
        };
      })
      .sort(
        (a, b) =>
          String(a.담당자).localeCompare(String(b.담당자), "ko-KR") ||
          String(a.마지막발주일).localeCompare(String(b.마지막발주일)) ||
          String(a.거래처명).localeCompare(String(b.거래처명), "ko-KR"),
      );
  };

  const downloadManagerLastOrders = (managers: string[]) => {
    if (!managers.length) {
      alert("담당자를 선택해주세요.");
      return;
    }
    const rows = buildManagerLastOrderRows(managers);
    exportExcel(rows, `담당자별_마지막발주일_${managers.join("_")}_${month}`);
  };

  const salesStatusExcelRows = sortedRows.map((r) => ({
    구분: r.label,
    채널: r.channel,
    전년동월: r.prevYearSales,
    "전년 Time gone": pct(r.prevYearTimeGoneGap),
    전년대비: pct(r.prevYearRate),
    전월: r.prevMonthSales,
    "전월 Time gone": pct(r.prevMonthTimeGoneGap),
    전월대비: pct(r.prevMonthRate),
    "당일까지 매출": r.currentSales,
    "당월 전체 매출": r.fullMonthSales,
    "TIME GONE 대비": pct(r.timeGoneGap),
    EST: r.est,
    "EST 달성률": pct(r.estRate),
    이익금액: r.profitAmount,
    이익률: pct(r.profitRate),
  }));

  const salesStatusColSpan = 1 + (showChannelColumn ? 1 : 0) + (!compact && isStoreListView ? 1 : 0) + 10;
  return (
    <>
      <div className="rounded-2xl border border-gray-300/70 bg-white/80 p-4 shadow-sm backdrop-blur">
        <div className="mb-3 flex flex-col gap-3 lg:flex-row lg:items-center lg:justify-between">
          <div>
            <h2 className="text-lg font-bold text-slate-900">매출현황</h2>
            <p className="mt-1 text-xs text-slate-500">
              상단 헤더를 클릭하면 내림차순/오름차순으로 정렬됩니다. 금액을
              클릭하면 해당 주문내역을 볼 수 있습니다.
            </p>
          </div>
          <div className="ml-auto flex flex-wrap items-center justify-end gap-2">
            {compact && (
              <select
                value={view}
                onChange={(e) => setView(e.target.value as SalesView)}
                className="w-[150px] rounded-lg border border-slate-300 bg-white px-3 py-2 text-xs font-semibold text-slate-700 shadow-sm outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
              >
                {SALES_VIEWS.map((v) => (
                  <option key={v} value={v}>
                    {v}
                  </option>
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
                  <option key={v} value={v}>
                    {v}
                  </option>
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
                  onChange={(e) =>
                    setOrderDateFilter(
                      e.target.value as "all" | "check7" | "no30",
                    )
                  }
                  className="w-[190px] rounded-lg border border-slate-300 bg-white/80 px-3 py-1.5 text-xs font-semibold outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                >
                  <option value="all">마지막 발주일 전체</option>
                  <option value="check7">확인필요 7일 이상</option>
                  <option value="no30">미발주 30일 이상</option>
                </select>
              )}
              <span className="text-xs font-medium text-slate-500">
                표시 기준: {view}
              </span>
              {view === "담당자별" && (
                <div className="flex max-w-full flex-wrap items-center gap-1 rounded-lg border border-slate-200 bg-slate-50 px-2 py-1">
                  <span className="mr-1 text-[12px] font-bold text-slate-600">
                    담당자 선택
                  </span>
                  {managerFilterOptions.map((manager) => (
                    <label
                      key={manager}
                      className="flex items-center gap-1 rounded-md bg-white px-2 py-1 text-[12px] font-semibold text-slate-700 ring-1 ring-slate-200"
                    >
                      <input
                        type="checkbox"
                        checked={selectedManagers.includes(manager)}
                        onChange={() => toggleSelectedManager(manager)}
                        className="h-3.5 w-3.5 rounded border-slate-300"
                      />
                      {manager}
                    </label>
                  ))}
                  {selectedManagers.length > 0 && (
                    <button
                      type="button"
                      onClick={() => setSelectedManagers([])}
                      className="rounded-md border border-slate-300 bg-white px-2 py-1 text-[12px] font-semibold text-slate-600 hover:bg-slate-50"
                    >
                      전체보기
                    </button>
                  )}
                </div>
              )}
              <div className="ml-auto flex flex-wrap items-center justify-end gap-2">
                {view === "담당자별" && (
                  <button
                    type="button"
                    onClick={() =>
                      downloadManagerLastOrders(
                        selectedManagers.length
                          ? selectedManagers
                          : managerFilterOptions,
                      )
                    }
                    className="rounded-lg bg-blue-600 px-3 py-1.5 text-xs font-semibold text-white shadow-sm hover:bg-blue-700"
                  >
                    선택 담당자 마지막 발주일 다운로드
                  </button>
                )}
                <button
                  type="button"
                  onClick={() => setInactiveOpen(true)}
                  className="rounded-lg border border-slate-300 bg-white px-3 py-1.5 text-xs font-semibold text-slate-700 shadow-sm hover:bg-slate-50"
                >
                  3개월 미주문 보기
                </button>
                <button
                  onClick={() =>
                    exportExcel(
                      salesStatusExcelRows,
                      `매출현황_${month}_${search || "전체"}_${view}`,
                    )
                  }
                  className="rounded-lg bg-emerald-600 px-3 py-1.5 text-xs font-semibold text-white shadow-sm hover:bg-emerald-700"
                >
                  매출현황 엑셀 다운로드
                </button>
              </div>
            </div>

            <div className="mb-3 grid grid-cols-2 gap-3 lg:grid-cols-5">
              <FilterAmountCard
                title="전년동월 매출"
                value={filteredPrevYearSales}
                tone="mint"
                onClick={() => openTotalDrill("prevYear")}
              />
              <FilterAmountCard
                title="전월 매출"
                value={filteredPrevMonthSales}
                tone="blue"
                onClick={() => openTotalDrill("prevMonth")}
              />
              <FilterAmountCard
                title="당일까지의 매출"
                value={filteredCurrentSales}
                tone="pink"
                onClick={() => openTotalDrill("current")}
              />
              <FilterAmountCard
                title="당월 발주 총 금액"
                value={filteredFullMonthSales}
                tone="pink"
                onClick={() => openTotalDrill("currentFullMonth")}
              />
              <FilterAmountCard
                title="이익 금액"
                value={filteredProfitAmount}
                tone="orange"
              />
            </div>
          </>
        )}

        <div className="relative max-h-[62vh] overflow-auto bg-white">
          <table
            className={`sales-status-table w-full ${compact ? "min-w-[1360px]" : "min-w-[1180px]"} table-fixed border-separate border-spacing-0 border border-gray-300 text-[12px] leading-tight`}
          >
            <thead>
              <tr className="bg-white">
                <ThCompactSortable
                  rowSpan={2}
                  w={isStoreListView ? (compact ? "w-[26%]" : "w-[19%]") : "w-[11%]"}
                  sortKey="label"
                  sortConfig={sortConfig}
                  onSort={requestSort}
                >
                  {isStoreListView ? "거래처" : view.replace("별", "")}
                </ThCompactSortable>
                {showChannelColumn && <ThCompact rowSpan={2} tone="gray" w="w-[7%]">채널</ThCompact>}
                {!compact && isStoreListView && (
                  <ThCompactSortable
                    rowSpan={2}
                    tone="gray"
                    w="w-[7%]"
                    sortKey="lastOrderDate"
                    sortConfig={sortConfig}
                    onSort={requestSort}
                  >
                    마지막발주일
                  </ThCompactSortable>
                )}
                <ThCompact colSpan={2} tone="mint" className="period-group-start">전년동월</ThCompact>
                <ThCompact colSpan={2} tone="blue" className="period-group-start">전월</ThCompact>
                <ThCompact colSpan={2} tone="pink" className="period-group-start">당월</ThCompact>
                <ThCompact colSpan={2} tone="yellow" className="period-group-start">EST</ThCompact>
                <ThCompact colSpan={2} tone="orange" className="period-group-start">이익</ThCompact>
              </tr>
              <tr>
                <ThCompactSortable right tone="mint" top="top-[31px]" className="period-group-start" w={compact ? "w-[6%]" : ""} sortKey="prevYearSales" sortConfig={sortConfig} onSort={requestSort}>매출</ThCompactSortable>
                <ThCompactSortable right tone="mint" top="top-[31px]" sortKey="prevYearTimeGoneGap" sortConfig={sortConfig} onSort={requestSort}>Time gone</ThCompactSortable>
                <ThCompactSortable right tone="blue" top="top-[31px]" className="period-group-start" w={compact ? "w-[6%]" : ""} sortKey="prevMonthSales" sortConfig={sortConfig} onSort={requestSort}>매출</ThCompactSortable>
                <ThCompactSortable right tone="blue" top="top-[31px]" sortKey="prevMonthTimeGoneGap" sortConfig={sortConfig} onSort={requestSort}>Time gone</ThCompactSortable>
                <ThCompactSortable right tone="pink" top="top-[31px]" className="period-group-start" sortKey="currentSales" sortConfig={sortConfig} onSort={requestSort}>당일까지 매출</ThCompactSortable>
                <ThCompactSortable right tone="pink" top="top-[31px]" sortKey="fullMonthSales" sortConfig={sortConfig} onSort={requestSort}>전체 매출</ThCompactSortable>
                <ThCompactSortable right tone="yellow" top="top-[31px]" className="period-group-start" w={compact ? "w-[5%]" : ""} sortKey="est" sortConfig={sortConfig} onSort={requestSort}>EST</ThCompactSortable>
                <ThCompactSortable right tone="yellow" top="top-[31px]" w={compact ? "w-[6%]" : ""} sortKey="estRate" sortConfig={sortConfig} onSort={requestSort}>EST 달성률</ThCompactSortable>
                <ThCompactSortable right tone="orange" top="top-[31px]" className="period-group-start" sortKey="profitAmount" sortConfig={sortConfig} onSort={requestSort}>이익금액</ThCompactSortable>
                <ThCompactSortable right tone="orange" top="top-[31px]" w="w-[5%]" sortKey="profitRate" sortConfig={sortConfig} onSort={requestSort}>이익률</ThCompactSortable>
              </tr>
            </thead>
            <tbody>
              {sortedRows.length === 0 ? (
                <tr>
                  <td
                    colSpan={salesStatusColSpan}
                    className="border p-8 text-center text-slate-500"
                  >
                    표시할 데이터가 없습니다.
                  </td>
                </tr>
              ) : (
                sortedRows.map((r) => {
                  const managerStoreRows =
                    view === "담당자별"
                      ? managerStoreRowsByKey.get(r.key) || []
                      : [];
                  return (
                    <Fragment key={r.key}>
                      <tr>
                        <TdCompact bold>{r.label}</TdCompact>
                        {showChannelColumn && <TdCompact bold>{r.channel}</TdCompact>}
                        {!compact && isStoreListView && (
                          <TdCompact>
                            <div className="text-center font-semibold text-slate-900 whitespace-nowrap">
                              {r.lastOrderDate}
                            </div>
                            <div
                              className={`mt-0.5 text-center text-[10px] font-semibold whitespace-nowrap ${r.daysSinceLastOrder >= 30 ? "text-red-600" : r.daysSinceLastOrder >= 7 ? "text-amber-600" : "text-slate-400"}`}
                            >
                              {r.lastOrderDate === "-"
                                ? "발주 없음"
                                : `${r.daysSinceLastOrder}일 경과`}
                            </div>
                          </TdCompact>
                        )}
                        <ClickableAmountCell
                          value={r.prevYearSales}
                          onClick={() => openDrill(r, "prevYear")}
                        />
                        <TdCompact right amount>
                          {pct(r.prevYearTimeGoneGap)}
                        </TdCompact>
                        <ClickableAmountCell
                          value={r.prevMonthSales}
                          onClick={() => openDrill(r, "prevMonth")}
                        />
                        <TdCompact right amount>
                          {pct(r.prevMonthTimeGoneGap)}
                        </TdCompact>
                        <ClickableAmountCell
                          value={r.currentSales}
                          onClick={() => openDrill(r, "current")}
                        />
                        <ClickableAmountCell
                          value={r.fullMonthSales}
                          onClick={() => openDrill(r, "currentFullMonth")}
                        />
                        <TdCompact right amount>
                          {won(r.est)}
                        </TdCompact>
                        <TdCompact right amount>
                          {pct(r.estRate)}
                        </TdCompact>
                        <TdCompact right amount>
                          {won(r.profitAmount)}
                        </TdCompact>
                        <TdCompact right amount>
                          {pct(r.profitRate)}
                        </TdCompact>
                      </tr>
                      {false &&
                        view === "담당자별" &&
                        managerStoreRows.length > 0 && (
                          <tr>
                            <td
                              colSpan={salesStatusColSpan}
                              className="border border-gray-300 bg-slate-50 px-3 py-2"
                            >
                              <div className="overflow-auto rounded-lg border border-gray-200 bg-white">
                                <table className="w-full min-w-[700px] border-separate border-spacing-0 text-[12px]">
                                  <thead>
                                    <tr className="bg-slate-100">
                                      <th className="border border-gray-200 px-2 py-1.5 text-center font-bold whitespace-nowrap">
                                        거래처코드
                                      </th>
                                      <th className="border border-gray-200 px-2 py-1.5 text-center font-bold whitespace-nowrap">
                                        거래처명
                                      </th>
                                      <th className="border border-gray-200 px-2 py-1.5 text-center font-bold whitespace-nowrap">
                                        전년동월
                                      </th>
                                      <th className="border border-gray-200 px-2 py-1.5 text-center font-bold whitespace-nowrap">
                                        전월
                                      </th>
                                      <th className="border border-gray-200 px-2 py-1.5 text-center font-bold whitespace-nowrap">
                                        당일까지 매출
                                      </th>
                                      <th className="border border-gray-200 px-2 py-1.5 text-center font-bold whitespace-nowrap">
                                        당월 전체 매출
                                      </th>
                                      <th className="border border-gray-200 px-2 py-1.5 text-center font-bold whitespace-nowrap">
                                        이익금액
                                      </th>
                                    </tr>
                                  </thead>
                                  <tbody>
                                    {managerStoreRows.map((storeRow) => (
                                      <tr
                                        key={`${r.key}-${storeRow.code}`}
                                        className="hover:bg-slate-50"
                                      >
                                        <td className="border border-gray-200 px-2 py-1.5 whitespace-nowrap">
                                          {storeRow.code}
                                        </td>
                                        <td className="border border-gray-200 px-2 py-1.5 font-semibold whitespace-nowrap">
                                          {storeRow.name}
                                        </td>
                                        <td className="border border-gray-200 px-2 py-1.5 text-right font-semibold whitespace-nowrap">
                                          {won(storeRow.prevYearSales)}
                                        </td>
                                        <td className="border border-gray-200 px-2 py-1.5 text-right font-semibold whitespace-nowrap">
                                          {won(storeRow.prevMonthSales)}
                                        </td>
                                        <td className="border border-gray-200 px-2 py-1.5 text-right font-semibold whitespace-nowrap">
                                          {won(storeRow.currentSales)}
                                        </td>
                                        <td className="border border-gray-200 px-2 py-1.5 text-right font-semibold whitespace-nowrap">
                                          {won(storeRow.fullMonthSales)}
                                        </td>
                                        <td className="border border-gray-200 px-2 py-1.5 text-right font-semibold whitespace-nowrap">
                                          {won(storeRow.profitAmount)}
                                        </td>
                                      </tr>
                                    ))}
                                  </tbody>
                                </table>
                              </div>
                            </td>
                          </tr>
                        )}
                    </Fragment>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>

      {drill && (
        <OrderDrillModal
          title={drill.title}
          rows={drill.rows}
          allSales={sales}
          onClose={() => setDrill(null)}
        />
      )}
      {inactiveOpen && (
        <InactiveOrdersModal
          stores={stores}
          sales={sales}
          month={month}
          onClose={() => setInactiveOpen(false)}
        />
      )}
    </>
  );
}

function ThCompactSortable({
  children,
  sortKey,
  sortConfig,
  onSort,
  right = false,
  w = "",
  tone = "default",
  rowSpan,
  top = "top-0",
  className = "",
}: {
  children: React.ReactNode;
  sortKey: SalesStatusSortKey;
  sortConfig: { key: SalesStatusSortKey; direction: SortDirection };
  onSort: (key: SalesStatusSortKey) => void;
  right?: boolean;
  w?: string;
  tone?: "default" | "mint" | "blue" | "pink" | "yellow" | "gray" | "purple" | "green" | "orange";
  rowSpan?: number;
  top?: string;
  className?: string;
}) {
  return (
    <ThCompact right={right} w={w} tone={tone} rowSpan={rowSpan} top={top} className={className}>
      <button
        type="button"
        onClick={() => onSort(sortKey)}
        className="flex w-full items-center justify-center gap-1 text-center"
      >
        <span>{children}</span>
        <span className="text-[11px] text-black">
          {sortArrow(sortConfig.key === sortKey, sortConfig.direction)}
        </span>
      </button>
    </ThCompact>
  );
}

function InactiveOrdersModal({
  sales,
  month,
  onClose,
}: {
  stores: Store[];
  sales: SalesRecord[];
  month: string;
  onClose: () => void;
}) {
  const startDate = threeMonthStart(month);
  const endDate = monthEnd(month);

  const itemMap = new Map<
    string,
    {
      itemCode: string;
      itemName: string;
      latest?: SalesRecord;
      recent: boolean;
    }
  >();
  sales.forEach((r) => {
    const key = `${r.itemCode}|${r.itemName}`;
    const current = itemMap.get(key) || {
      itemCode: r.itemCode || "-",
      itemName: r.itemName || "미지정",
      latest: undefined,
      recent: false,
    };
    if (inRange(r.saleDate, startDate, endDate)) current.recent = true;
    if (!current.latest || r.saleDate > current.latest.saleDate)
      current.latest = r;
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
    .sort(
      (a, b) =>
        String(a.lastDate).localeCompare(String(b.lastDate)) ||
        a.itemName.localeCompare(b.itemName, "ko-KR"),
    );

  const itemExcelRows = inactiveItems.map((r) => ({
    상품코드: r.itemCode,
    상품명: r.itemName,
    마지막주문일: r.lastDate,
    마지막거래처: r.lastStore,
    마지막매출금액: r.lastAmount,
  }));

  function sendInactiveMail() {
    const to = window.prompt(
      "메일을 받을 주소를 입력하세요. 여러 명이면 쉼표로 구분해주세요.",
    );
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
        <div className="flex flex-col gap-3 border-b border-gray-300 p-4 lg:flex-row lg:items-center lg:justify-between">
          <div>
            <h3 className="text-lg font-bold text-slate-900">
              3개월 미주문 품목 현황
            </h3>
            <p className="mt-1 text-xs text-slate-500">
              기준 기간: {startDate} ~ {endDate} · 총{" "}
              {inactiveItems.length.toLocaleString("ko-KR")}건
            </p>
          </div>
          <div className="flex flex-wrap gap-2">
            <button
              type="button"
              onClick={sendInactiveMail}
              className="rounded-lg bg-blue-600 px-3 py-2 text-xs font-semibold text-white shadow-sm hover:bg-blue-700"
            >
              메일 전송(파일 다운로드)
            </button>
            <button
              type="button"
              onClick={() =>
                exportExcel(itemExcelRows, `3개월_미주문_품목별_${month}`)
              }
              className="rounded-lg bg-emerald-600 px-3 py-2 text-xs font-semibold text-white shadow-sm hover:bg-emerald-700"
            >
              엑셀 다운로드
            </button>
            <button
              type="button"
              onClick={onClose}
              className="rounded-lg border border-slate-300 bg-white px-3 py-2 text-xs font-semibold text-slate-700 hover:bg-slate-50"
            >
              닫기
            </button>
          </div>
        </div>

        <div className="min-h-0 flex-1 overflow-auto bg-white px-4 pb-4 pt-0">
          <table className="w-full min-w-[860px] border-separate border-spacing-0 border border-gray-300 bg-white text-xs">
            <thead>
              <tr className="bg-slate-100">
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  상품코드
                </th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  상품명
                </th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  마지막 주문일
                </th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  마지막 거래처
                </th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  마지막 매출금액
                </th>
              </tr>
            </thead>
            <tbody>
              {inactiveItems.length === 0 ? (
                <tr>
                  <td
                    colSpan={5}
                    className="border p-8 text-center text-slate-500"
                  >
                    3개월 미주문 품목이 없습니다.
                  </td>
                </tr>
              ) : (
                inactiveItems.map((r) => (
                  <tr
                    key={`${r.itemCode}|${r.itemName}`}
                    className="hover:bg-slate-50"
                  >
                    <td className="border px-2 py-2">{r.itemCode}</td>
                    <td className="border px-2 py-2 font-semibold">
                      {r.itemName}
                    </td>
                    <td className="border px-2 py-2">{r.lastDate}</td>
                    <td className="border px-2 py-2">{r.lastStore}</td>
                    <td className="border px-2 py-2 text-right font-semibold text-slate-900">
                      {won(r.lastAmount)}
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function DormantAccountPage({
  stores,
  sales,
  month,
}: {
  stores: Store[];
  sales: SalesRecord[];
  month: string;
}) {
  const [tab, setTab] = useState<InactiveOrderTab>("거래처별");
  const startDate = threeMonthStart(month);
  const endDate = monthEnd(month);
  const activeStores = stores.filter((s) => s.status === "거래중");
  const recentSales = sales.filter((s) =>
    inRange(s.saleDate, startDate, endDate),
  );
  const recentStoreCodes = new Set(recentSales.map((s) => s.storeCode));
  const latestSaleByStore = new Map<string, SalesRecord>();

  sales.forEach((r) => {
    const prev = latestSaleByStore.get(r.storeCode);
    if (!prev || r.saleDate > prev.saleDate)
      latestSaleByStore.set(r.storeCode, r);
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
    .sort(
      (a, b) =>
        String(a.lastDate).localeCompare(String(b.lastDate)) ||
        a.name.localeCompare(b.name, "ko-KR"),
    );

  const itemMap = new Map<
    string,
    {
      itemCode: string;
      itemName: string;
      latest?: SalesRecord;
      recent: boolean;
    }
  >();
  sales.forEach((r) => {
    const key = `${r.itemCode}|${r.itemName}`;
    const current = itemMap.get(key) || {
      itemCode: r.itemCode || "-",
      itemName: r.itemName || "미지정",
      latest: undefined,
      recent: false,
    };
    if (inRange(r.saleDate, startDate, endDate)) current.recent = true;
    if (!current.latest || r.saleDate > current.latest.saleDate)
      current.latest = r;
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
    .sort(
      (a, b) =>
        String(a.lastDate).localeCompare(String(b.lastDate)) ||
        a.itemName.localeCompare(b.itemName, "ko-KR"),
    );

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
    const to = window.prompt(
      "메일을 받을 주소를 입력하세요. 여러 명이면 쉼표로 구분해주세요.",
    );
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
    <div className="rounded-xl border border-gray-300/70 bg-white/80 p-4 shadow-sm backdrop-blur">
      <div className="mb-3 flex flex-col gap-3 lg:flex-row lg:items-center lg:justify-between">
        <div>
          <h2 className="text-base font-bold text-slate-900">휴면거래처관리</h2>
          <p className="mt-1 text-xs text-slate-500">
            3개월 미주문 거래처와 품목을 확인하고 엑셀 다운로드 또는 메일 전송용
            파일을 만들 수 있습니다. 기준 기간: {startDate} ~ {endDate}
          </p>
        </div>
        <div className="flex flex-wrap gap-2">
          <button
            type="button"
            onClick={sendInactiveMail}
            className="rounded-lg bg-blue-600 px-3 py-2 text-xs font-semibold text-white shadow-sm hover:bg-blue-700"
          >
            메일 전송용 파일 다운로드
          </button>
          <button
            type="button"
            onClick={() =>
              exportExcel(
                tab === "거래처별" ? storeExcelRows : itemExcelRows,
                `3개월_미주문_${tab}_${month}`,
              )
            }
            className="rounded-lg bg-emerald-600 px-3 py-2 text-xs font-semibold text-white shadow-sm hover:bg-emerald-700"
          >
            엑셀 다운로드
          </button>
        </div>
      </div>

      <div className="mb-3 flex gap-2 border-b border-gray-300 pt-1">
        {(["거래처별", "품목별"] as InactiveOrderTab[]).map((item) => (
          <button
            key={item}
            type="button"
            onClick={() => setTab(item)}
            className={`rounded-t-lg px-4 py-2 text-sm font-semibold ${tab === item ? "bg-blue-600 text-white" : "bg-slate-100 text-slate-700 hover:bg-slate-200"}`}
          >
            {item}{" "}
            {item === "거래처별" ? inactiveStores.length : inactiveItems.length}
            건
          </button>
        ))}
      </div>

      <div className="max-h-[62vh] overflow-auto">
        {tab === "거래처별" ? (
          <table className="w-full min-w-[920px] border border-gray-300 text-xs">
            <thead>
              <tr className="bg-slate-100">
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  거래처코드
                </th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  거래처명
                </th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  브랜드
                </th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  담당자
                </th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  채널
                </th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  마지막 주문일
                </th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  마지막 매출금액
                </th>
              </tr>
            </thead>
            <tbody>
              {inactiveStores.length === 0 ? (
                <tr>
                  <td
                    colSpan={7}
                    className="border p-8 text-center text-slate-500"
                  >
                    3개월 미주문 거래처가 없습니다.
                  </td>
                </tr>
              ) : (
                inactiveStores.map((r) => (
                  <tr key={r.code} className="hover:bg-slate-50">
                    <td className="border px-2 py-2">{r.code}</td>
                    <td className="border px-2 py-2 font-semibold">{r.name}</td>
                    <td className="border px-2 py-2">{r.brand}</td>
                    <td className="border px-2 py-2">{r.manager}</td>
                    <td className="border px-2 py-2">{r.channel}</td>
                    <td className="border px-2 py-2">{r.lastDate}</td>
                    <td className="border px-2 py-2 text-right font-semibold text-slate-900">
                      {won(r.lastAmount)}
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        ) : (
          <table className="w-full min-w-[860px] border-separate border-spacing-0 border border-gray-300 bg-white text-xs">
            <thead>
              <tr className="bg-slate-100">
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  상품코드
                </th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  상품명
                </th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  마지막 주문일
                </th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  마지막 거래처
                </th>
                <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-2 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  마지막 매출금액
                </th>
              </tr>
            </thead>
            <tbody>
              {inactiveItems.length === 0 ? (
                <tr>
                  <td
                    colSpan={5}
                    className="border p-8 text-center text-slate-500"
                  >
                    3개월 미주문 품목이 없습니다.
                  </td>
                </tr>
              ) : (
                inactiveItems.map((r) => (
                  <tr
                    key={`${r.itemCode}|${r.itemName}`}
                    className="hover:bg-slate-50"
                  >
                    <td className="border px-2 py-2">{r.itemCode}</td>
                    <td className="border px-2 py-2 font-semibold">
                      {r.itemName}
                    </td>
                    <td className="border px-2 py-2">{r.lastDate}</td>
                    <td className="border px-2 py-2">{r.lastStore}</td>
                    <td className="border px-2 py-2 text-right font-semibold text-slate-900">
                      {won(r.lastAmount)}
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}

function ClickableAmountCell({
  value,
  onClick,
}: {
  value: number;
  onClick: () => void;
}) {
  const disabled = !value;
  return (
    <td className="border border-gray-300 bg-white px-1.5 py-2 text-right align-middle whitespace-nowrap break-keep">
      <button
        type="button"
        onClick={onClick}
        disabled={disabled}
        className={`w-full whitespace-nowrap text-right text-[12px] font-bold leading-snug underline-offset-2 ${disabled ? "cursor-default text-slate-400" : "text-slate-900 hover:underline"}`}
        title={disabled ? "주문내역이 없습니다." : "주문내역 보기"}
      >
        {won(value)}
      </button>
    </td>
  );
}

function compactList(values: string[]) {
  const unique = Array.from(
    new Set(values.map((v) => norm(v) || "미지정")),
  ).filter(Boolean);
  if (!unique.length) return "-";
  if (unique.length <= 3) return unique.join(", ");
  return `${unique.slice(0, 3).join(", ")} 외 ${unique.length - 3}개`;
}

function OrderDrillModal({
  title,
  rows,
  allSales,
  onClose,
}: {
  title: string;
  rows: SalesRecord[];
  allSales: SalesRecord[];
  onClose: () => void;
}) {
  const [itemDrill, setItemDrill] = useState<{
    itemCode: string;
    itemName: string;
    rows: SalesRecord[];
  } | null>(null);
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
      .sort(
        (a, b) =>
          b.saleDate.localeCompare(a.saleDate) ||
          a.storeName.localeCompare(b.storeName, "ko-KR"),
      );
    setItemDrill({ itemCode, itemName, rows: filteredRows });
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/55 p-3">
      <style jsx global>{`
        .order-popup-table,
        .order-popup-table th,
        .order-popup-table td {
          border-color: #d1d5db !important;
        }
      `}</style>
      <div className="flex max-h-[94vh] w-full max-w-[96vw] flex-col rounded-2xl bg-white shadow-2xl">
        <div className="sticky top-0 z-20 flex flex-col gap-3 border-b border-gray-300 bg-white p-5 shadow-sm xl:flex-row xl:items-start xl:justify-between">
          <div className="min-w-0 flex-1">
            <div className="flex flex-wrap items-center gap-2">
              <h3 className="text-xl font-bold text-slate-900">
                {title} 주문내역
              </h3>
              <span className="rounded-full bg-slate-100 px-3 py-1 text-xs font-semibold text-slate-700">
                브랜드: {summaryBrand}
              </span>
              <span className="rounded-full bg-slate-100 px-3 py-1 text-xs font-semibold text-slate-700">
                담당자: {summaryManager}
              </span>
              <span className="rounded-full bg-slate-100 px-3 py-1 text-xs font-semibold text-slate-700">
                채널: {summaryChannel}
              </span>
            </div>
            <p className="mt-2 text-sm text-slate-500">
              총 {rows.length.toLocaleString("ko-KR")}건 · 수량{" "}
              {won(totalQuantity)} · 매출 {won(totalSales)}원 · 이익{" "}
              {won(totalProfit)}원
            </p>
            <p className="mt-1 text-xs text-slate-400">
              상품코드를 클릭하면 해당 품목이 나간 전체 거래처 주문내역을 볼 수
              있습니다.
            </p>
          </div>
          <div className="flex shrink-0 gap-2">
            <button
              type="button"
              onClick={() =>
                exportExcel(
                  excelRows,
                  `${title.replaceAll(" ", "_").replaceAll("·", "_")}_주문내역`,
                )
              }
              className="rounded-lg bg-emerald-600 px-4 py-2 text-sm font-semibold text-white shadow-sm hover:bg-emerald-700"
            >
              주문내역 엑셀 다운로드
            </button>
            <button
              type="button"
              onClick={onClose}
              className="rounded-lg border border-gray-300 bg-white px-4 py-2 text-sm font-semibold text-slate-700 hover:bg-slate-50"
            >
              닫기
            </button>
          </div>
        </div>

        <div className="relative isolate min-h-0 flex-1 overflow-auto bg-white px-5 pb-5 pt-0">
          <table className="order-popup-table w-full min-w-[1100px] border-separate border-spacing-0 border border-gray-300 bg-white text-sm">
            <thead className="sticky top-0 z-[80] bg-white shadow-[0_2px_0_0_#e2e8f0]">
              <tr className="bg-white">
                <th className="sticky top-0 z-[80] border border-gray-300 bg-white px-3 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  주문일
                </th>
                <th className="sticky top-0 z-[80] border border-gray-300 bg-white px-3 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  거래처
                </th>
                <th className="sticky top-0 z-[80] border border-gray-300 bg-white px-3 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  상품코드
                </th>
                <th className="sticky top-0 z-[80] border border-gray-300 bg-white px-3 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  상품명
                </th>
                <th className="sticky top-0 z-[80] border border-gray-300 bg-white px-3 py-2 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  수량
                </th>
                <th className="sticky top-0 z-[80] border border-gray-300 bg-white px-3 py-2 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  매출금액
                </th>
                <th className="sticky top-0 z-[80] border border-gray-300 bg-white px-3 py-2 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  원가금액
                </th>
                <th className="sticky top-0 z-[80] border border-gray-300 bg-white px-3 py-2 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  이익금액
                </th>
                <th className="sticky top-0 z-[80] border border-gray-300 bg-white px-3 py-2 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  이익률
                </th>
              </tr>
            </thead>
            <tbody>
              {rows.length === 0 ? (
                <tr>
                  <td
                    colSpan={9}
                    className="border border-gray-300 p-10 text-center text-slate-500"
                  >
                    표시할 주문내역이 없습니다.
                  </td>
                </tr>
              ) : (
                rows.map((r) => (
                  <tr key={r.id} className="hover:bg-slate-50">
                    <td className="border border-gray-300 px-3 py-2">
                      {r.saleDate}
                    </td>
                    <td className="border border-gray-300 px-3 py-2 font-semibold">
                      {r.storeName}
                    </td>
                    <td className="border border-gray-300 px-3 py-2">
                      <button
                        type="button"
                        onClick={() => openItemDrill(r.itemCode, r.itemName)}
                        className="font-bold text-slate-900 underline-offset-2 hover:underline"
                        title="품목별 주문내역 보기"
                      >
                        {r.itemCode}
                      </button>
                    </td>
                    <td className="border border-gray-300 px-3 py-2">
                      {r.itemName}
                    </td>
                    <td className="border border-gray-300 px-3 py-2 text-right">
                      {won(r.quantity)}
                    </td>
                    <td className="border border-gray-300 px-3 py-2 text-right text-base font-bold text-slate-900">
                      {won(r.salesAmount)}
                    </td>
                    <td className="border border-gray-300 px-3 py-2 text-right font-semibold text-slate-900">
                      {won(r.costAmount)}
                    </td>
                    <td className="border border-gray-300 px-3 py-2 text-right font-semibold text-slate-900">
                      {won(r.profitAmount)}
                    </td>
                    <td className="border border-gray-300 px-3 py-2 text-right text-slate-900">
                      {pct(r.profitRate)}
                    </td>
                  </tr>
                ))
              )}
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

function ItemDrillModal({
  itemCode,
  itemName,
  rows,
  onClose,
}: {
  itemCode: string;
  itemName: string;
  rows: SalesRecord[];
  onClose: () => void;
}) {
  const totalSales = sum(rows, "salesAmount");
  const totalQuantity = sum(rows, "quantity");
  const totalCost = sum(rows, "costAmount");
  const totalProfit = sum(rows, "profitAmount");
  const totalProfitRate = weightedProfitRate(rows);
  const excelRows = orderRowsForExcel(rows);

  return (
    <div className="absolute inset-0 z-[60] flex items-center justify-center bg-slate-950/45 p-4">
      <style jsx global>{`
        .order-popup-table,
        .order-popup-table th,
        .order-popup-table td {
          border-color: #d1d5db !important;
        }
      `}</style>
      <div className="flex max-h-[88vh] w-full max-w-[90vw] flex-col rounded-2xl bg-white shadow-2xl ring-1 ring-slate-200">
        <div className="sticky top-0 z-20 flex flex-col gap-3 border-b border-gray-300 bg-white p-5 shadow-sm xl:flex-row xl:items-start xl:justify-between">
          <div className="min-w-0 flex-1">
            <h3 className="text-xl font-bold text-slate-900">
              품목별 전체 거래처 주문내역
            </h3>
            <p className="mt-1 text-sm font-semibold text-slate-700">
              {itemCode} · {itemName}
            </p>
            <p className="mt-2 text-sm text-slate-500">
              총 {rows.length.toLocaleString("ko-KR")}건 · 수량{" "}
              {won(totalQuantity)} · 매출 {won(totalSales)}원 · 원가{" "}
              {won(totalCost)}원 · 이익 {won(totalProfit)}원 · 이익률{" "}
              {pct(totalProfitRate)}
            </p>
          </div>
          <div className="flex shrink-0 gap-2">
            <button
              type="button"
              onClick={() =>
                exportExcel(
                  excelRows,
                  `${itemCode}_${itemName}_품목별_전체거래처_주문내역`,
                )
              }
              className="rounded-lg bg-emerald-600 px-4 py-2 text-sm font-semibold text-white shadow-sm hover:bg-emerald-700"
            >
              품목별 전체거래처 엑셀 다운로드
            </button>
            <button
              type="button"
              onClick={onClose}
              className="rounded-lg border border-gray-300 bg-white px-4 py-2 text-sm font-semibold text-slate-700 hover:bg-slate-50"
            >
              닫기
            </button>
          </div>
        </div>

        <div className="relative isolate min-h-0 flex-1 overflow-auto bg-white px-5 pb-5 pt-0">
          <table className="order-popup-table w-full min-w-[920px] border-separate border-spacing-0 border border-gray-300 bg-white text-sm">
            <thead className="sticky top-0 z-[80] bg-white shadow-[0_2px_0_0_#e2e8f0]">
              <tr className="bg-white">
                <th className="sticky top-0 z-[80] border border-gray-300 bg-white px-3 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  주문일
                </th>
                <th className="sticky top-0 z-[80] border border-gray-300 bg-white px-3 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  거래처
                </th>
                <th className="sticky top-0 z-[80] border border-gray-300 bg-white px-3 py-2 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  담당자
                </th>
                <th className="sticky top-0 z-[80] border border-gray-300 bg-white px-3 py-2 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  수량
                </th>
                <th className="sticky top-0 z-[80] border border-gray-300 bg-white px-3 py-2 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  매출금액
                </th>
                <th className="sticky top-0 z-[80] border border-gray-300 bg-white px-3 py-2 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  원가금액
                </th>
                <th className="sticky top-0 z-[80] border border-gray-300 bg-white px-3 py-2 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  이익금액
                </th>
                <th className="sticky top-0 z-[80] border border-gray-300 bg-white px-3 py-2 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">
                  이익률
                </th>
              </tr>
            </thead>
            <tbody>
              {rows.length === 0 ? (
                <tr>
                  <td
                    colSpan={8}
                    className="border border-gray-300 p-10 text-center text-slate-500"
                  >
                    표시할 품목별 주문내역이 없습니다.
                  </td>
                </tr>
              ) : (
                rows.map((r) => (
                  <tr key={r.id} className="hover:bg-slate-50">
                    <td className="border border-gray-300 px-3 py-2">
                      {r.saleDate}
                    </td>
                    <td className="border border-gray-300 px-3 py-2 font-semibold">
                      {r.storeName}
                    </td>
                    <td className="border border-gray-300 px-3 py-2">
                      {r.manager || "미지정"}
                    </td>
                    <td className="border border-gray-300 px-3 py-2 text-right">
                      {won(r.quantity)}
                    </td>
                    <td className="border border-gray-300 px-3 py-2 text-right text-base font-bold text-slate-900">
                      {won(r.salesAmount)}
                    </td>
                    <td className="border border-gray-300 px-3 py-2 text-right font-semibold text-slate-900">
                      {won(r.costAmount)}
                    </td>
                    <td className="border border-gray-300 px-3 py-2 text-right font-semibold text-slate-900">
                      {won(r.profitAmount)}
                    </td>
                    <td className="border border-gray-300 px-3 py-2 text-right text-slate-900">
                      {pct(r.profitRate)}
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function FilterAmountCard({
  title,
  value,
  tone,
  onClick,
}: {
  title: string;
  value: number | string;
  tone: "mint" | "blue" | "pink" | "yellow" | "green" | "orange";
  onClick?: () => void;
}) {
  const toneClass =
    tone === "mint"
      ? "border-lime-300 bg-lime-100"
      : tone === "blue"
        ? "border-sky-300 bg-sky-100"
        : tone === "pink"
          ? "border-pink-300 bg-pink-100"
          : tone === "yellow"
            ? "border-yellow-300 bg-yellow-100"
            : tone === "orange"
              ? "border-orange-300 bg-orange-100"
              : "border-white bg-white";

  return (
    <button
      type="button"
      onClick={onClick}
      disabled={!onClick}
      className={`rounded-xl border px-3 py-2 text-left shadow-sm ${toneClass} ${onClick ? "cursor-pointer hover:shadow-md" : "cursor-default"}`}
    >
      <p className="text-[13px] font-bold text-black">{title}</p>
      <p className="mt-1 break-all text-right text-[18px] font-bold text-black">
        {typeof value === "number" ? won(value) : value}
      </p>
    </button>
  );
}

function ThCompact({
  children,
  right = false,
  w = "",
  tone = "default",
  rowSpan,
  colSpan,
  top = "top-0",
  className = "",
}: {
  children: React.ReactNode;
  right?: boolean;
  w?: string;
  tone?: "default" | "mint" | "blue" | "pink" | "yellow" | "gray" | "purple" | "green" | "orange";
  rowSpan?: number;
  colSpan?: number;
  top?: string;
  className?: string;
}) {
  const toneClass =
    tone === "mint"
      ? "border-[#E5E7EB] bg-[#F7FCEB] text-black"
      : tone === "blue"
        ? "border-[#E5E7EB] bg-[#F3FAFD] text-black"
        : tone === "pink"
          ? "border-[#E5E7EB] bg-[#FFF7FA] text-black"
          : tone === "yellow"
            ? "border-[#E5E7EB] bg-[#FFFDF2] text-black"
            : tone === "orange"
              ? "border-[#E5E7EB] bg-[#FFF9F3] text-black"
              : "border-[#E5E7EB] bg-white text-black";

  return (
    <th
      rowSpan={rowSpan}
      colSpan={colSpan}
      className={`sticky ${top} z-50 border px-0.5 py-1.5 align-middle text-center text-[14px] font-bold leading-tight whitespace-nowrap break-keep bg-clip-padding ${toneClass} ${w} ${className}`}
    >
      {children}
    </th>
  );
}

function TdCompact({
  children,
  right = false,
  bold = false,
  color = "",
  amount = false,
}: {
  children: React.ReactNode;
  right?: boolean;
  bold?: boolean;
  color?: string;
  amount?: boolean;
}) {
  return (
    <td
      className={`border border-gray-300 bg-white px-0.5 py-1.5 align-middle text-center whitespace-nowrap break-keep ${bold ? "font-semibold" : ""} ${amount ? "text-[11px] font-bold leading-snug text-slate-900" : ""} ${color}`}
    >
      {children}
    </td>
  );
}

function SalesCompare({
  stores,
  sales,
  month,
  date,
}: {
  stores: Store[];
  sales: SalesRecord[];
  month: string;
  date: string;
}) {
  const [search, setSearch] = useState("");
  const [searchDraft, setSearchDraft] = useState("");
  const stMap = storeMap(stores);
  const storeCodes = Array.from(
    new Set([...stores.map((s) => s.code), ...sales.map((s) => s.storeCode)]),
  );
  const rows = storeCodes
    .map((code) => {
      const store = stMap.get(code);
      const current = sales.filter(
        (s) =>
          s.period === "current" &&
          s.storeCode === code &&
          inRange(s.saleDate, monthStart(month), date),
      );
      const prevMonth = sales.filter(
        (s) =>
          s.period === "prevMonth" &&
          s.refMonth === month &&
          s.storeCode === code,
      );
      const prevYear = sales.filter(
        (s) =>
          s.period === "prevYear" &&
          s.refMonth === month &&
          s.storeCode === code,
      );
      const currentSales = sum(current, "salesAmount");
      const prevMonthSales = sum(prevMonth, "salesAmount");
      const prevYearSales = sum(prevYear, "salesAmount");
      return {
        code,
        name:
          store?.name ||
          current[0]?.storeName ||
          prevMonth[0]?.storeName ||
          prevYear[0]?.storeName ||
          "",
        currentSales,
        prevMonthSales,
        prevYearSales,
        prevMonthRate: prevMonthSales
          ? ((currentSales - prevMonthSales) / prevMonthSales) * 100
          : 0,
        prevYearRate: prevYearSales
          ? ((currentSales - prevYearSales) / prevYearSales) * 100
          : 0,
      };
    })
    .filter((r) =>
      `${r.code} ${r.name}`.toLowerCase().includes(search.toLowerCase()),
    );

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
    <div className="rounded-xl border border-gray-300/70 bg-white/80 p-4 shadow-sm backdrop-blur">
      <div className="mb-3 flex flex-col gap-2 lg:flex-row lg:items-center lg:justify-between">
        <div>
          <h2 className="text-lg font-bold text-slate-900">매출비교</h2>
          <p className="mt-1 text-sm text-slate-500">
            거래처별 전년동월 / 전월 / 당월 매출을 비교합니다.
          </p>
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
            onClick={() =>
              exportExcel(salesCompareExcelRows, `매출비교_${month}`)
            }
            className="rounded-xl bg-emerald-600 px-4 py-2 text-sm font-semibold text-white shadow-sm hover:bg-emerald-700"
          >
            매출비교 엑셀 다운로드
          </button>
        </div>
      </div>

      <div className="max-h-[68vh] overflow-visible">
        <table className="w-full table-fixed border border-gray-300 text-xs">
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
  salesActions,
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
  itemMasters,
  setItemMasters,
}: {
  stores: Store[];
  setStores: (v: Store[]) => void;
  sales: SalesRecord[];
  setSales: React.Dispatch<React.SetStateAction<SalesRecord[]>>;
  salesActions: SalesStorageActions;
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
  itemMasters: ItemMasterRecord[];
  setItemMasters: (v: ItemMasterRecord[]) => void;
}) {
  const [tab, setTab] = useState<MonthStartTab>("거래처 리스트");

  return (
    <div className="space-y-3">
      <div className="border-b border-slate-300 bg-white px-2">
        <div className="flex min-w-max gap-1 overflow-x-auto">
          {MONTH_TABS.map((t) => (
            <button
              key={t}
              onClick={() => setTab(t)}
              className={`border-b-2 px-4 py-2.5 text-sm font-bold transition ${
                tab === t
                  ? "border-blue-600 text-blue-700"
                  : "border-transparent text-slate-600 hover:border-slate-300 hover:text-slate-900"
              }`}
            >
              {t}
            </button>
          ))}
        </div>
      </div>

      {tab === "거래처 리스트" && (
        <StoreListManagement
          section="list"
          stores={stores}
          setStores={setStores}
          sales={sales}
          month={month}
          codeMappings={codeMappings}
        />
      )}
      {tab === "기준정보" && (
        <StoreListManagement
          section="reference"
          stores={stores}
          setStores={setStores}
          sales={sales}
          month={month}
          codeMappings={codeMappings}
        />
      )}
      {tab === "업로드 관리" && (
        <UploadPage
          stores={stores}
          setStores={setStores}
          sales={sales}
          setSales={setSales}
          salesActions={salesActions}
          month={month}
          date={date}
          timeConfigs={timeConfigs}
          setTimeConfigs={setTimeConfigs}
          itemMasters={itemMasters}
          setItemMasters={setItemMasters}
        />
      )}
      {tab === "이익금액 검증표" && (
        <ProfitValidationPanel sales={sales} month={month} date={date} />
      )}
    </div>
  );
}

function StoreListManagement({
  section,
  stores,
  setStores,
  sales,
  month,
  codeMappings,
}: {
  section: "list" | "reference";
  stores: Store[];
  setStores: (v: Store[]) => void;
  sales: SalesRecord[];
  month: string;
  codeMappings: StoreCodeMapping[];
}) {
  type ListTab = "기존거래처 리스트" | "전년동월 리스트" | "총 거래처 리스트" | "기타 관리";
  const [listTab, setListTab] = useState<ListTab>("기존거래처 리스트");
  const [search, setSearch] = useState("");
  const [otherTab, setOtherTab] = useState<"담당자 관리" | "브랜드 관리" | "채널 관리">("담당자 관리");
  const [channelTab, setChannelTab] = useState<"채널 1" | "채널 2">("채널 1");
  const [selectedCodes, setSelectedCodes] = useState<Set<string>>(new Set());
  const [bulkManager, setBulkManager] = useState("");
  const [bulkBrand, setBulkBrand] = useState("");
  const [bulkChannel1, setBulkChannel1] = useState("");
  const [bulkChannel2, setBulkChannel2] = useState<"" | "매장" | "비매장">("");
  const [newManager, setNewManager] = useState("");
  const [newChannel1, setNewChannel1] = useState("");
  const [showNewOnly, setShowNewOnly] = useState(false);
  const [savedManagers, setSavedManagers] = useLocal<string[]>(
    "month-start-manager-options-v1",
    ["SY", "KT", "SW", "NH", "BOMI", "BM"],
  );
  const [deletedManagers, setDeletedManagers] = useLocal<string[]>(
    "month-start-deleted-manager-options-v1",
    [],
  );
  const [savedChannel1Options, setSavedChannel1Options] = useLocal<string[]>(
    "month-start-channel1-options-v1",
    ["도매", "체인", "권역배송", "온라인", "식자재마트"],
  );
  const normalizedSearch = search.trim().toLowerCase();

  useEffect(() => {
    setListTab(section === "list" ? "기존거래처 리스트" : "기타 관리");
    setSearch("");
    setSelectedCodes(new Set());
    setShowNewOnly(false);
  }, [section]);

  type SalesStoreSummary = {
    code: string;
    name: string;
    channel: string;
    manager: string;
    storeType: string;
    brand: string;
    amount: number;
  };

  const summarize = (rows: SalesRecord[]) => {
    const map = new Map<string, SalesStoreSummary>();
    rows.forEach((row) => {
      const key = `${norm(row.storeCode)}|${normalizeStoreNameKey(row.storeName)}`;
      const saved = stores.find((store) => store.code === row.storeCode);
      const previous = map.get(key) || {
        code: row.storeCode || "-",
        name: row.storeName || row.storeCode || "미지정",
        channel: saved?.channel || row.channel || "미지정",
        manager: saved?.manager || row.manager || "",
        storeType: saved?.storeType || row.storeType || "비매장",
        brand: displayBrand(saved?.brand || row.brand),
        amount: 0,
      };
      previous.amount += Number(row.salesAmount || 0);
      map.set(key, previous);
    });
    return Array.from(map.values());
  };

  const currentRows = useMemo(
    () =>
      summarize(
        sales.filter(
          (row) =>
            row.period === "current" &&
            inRange(row.saleDate, monthStart(month), monthEnd(month)),
        ),
      ),
    [sales, stores, month],
  );

  const prevMonthRows = useMemo(
    () =>
      summarize(
        sales.filter(
          (row) => row.period === "prevMonth" && row.refMonth === month,
        ),
      ),
    [sales, stores, month],
  );

  const prevYearRows = useMemo(
    () =>
      summarize(
        sales.filter(
          (row) => row.period === "prevYear" && row.refMonth === month,
        ),
      ),
    [sales, stores, month],
  );

  const existingRows = useMemo(() => {
    const currentByCode = new Map(currentRows.map((row) => [norm(row.code), row]));
    const currentByName = new Map(
      currentRows.map((row) => [normalizeStoreNameKey(row.name), row]),
    );
    const prevByCode = new Map(prevMonthRows.map((row) => [norm(row.code), row]));
    const prevByName = new Map(
      prevMonthRows.map((row) => [normalizeStoreNameKey(row.name), row]),
    );
    const keys = new Set<string>();
    currentRows.forEach((row) => keys.add(`C|${row.code}|${row.name}`));
    prevMonthRows.forEach((row) => keys.add(`P|${row.code}|${row.name}`));

    const result = new Map<string, SalesStoreSummary & {
      currentAmount: number;
      prevAmount: number;
      statusLabel: string;
      note: string;
    }>();

    [...currentRows, ...prevMonthRows].forEach((source) => {
      const currentBySameCode = currentByCode.get(norm(source.code));
      const currentBySameName = currentByName.get(normalizeStoreNameKey(source.name));
      const prevBySameCode = prevByCode.get(norm(source.code));
      const prevBySameName = prevByName.get(normalizeStoreNameKey(source.name));
      const current = currentBySameCode || currentBySameName;
      const prev = prevBySameCode || prevBySameName;
      const canonical = current || prev || source;
      const key = current?.code || prev?.code || source.code || source.name;

      let statusLabel = "동일";
      let note = "당월과 전월의 사업자번호와 거래처명이 같습니다.";
      if (current && !prev) {
        statusLabel = "당월 신규";
        note = "당월에만 존재하는 거래처입니다.";
      } else if (!current && prev) {
        statusLabel = "당월 미거래";
        note = "전월에는 있었지만 당월 매출에는 없습니다.";
      } else if (current && prev) {
        if (norm(current.code) !== norm(prev.code)) {
          statusLabel = "사업자번호 변경";
          note = `전월 ${prev.code} → 당월 ${current.code}`;
        } else if (
          normalizeStoreNameKey(current.name) !== normalizeStoreNameKey(prev.name)
        ) {
          statusLabel = "거래처명 변경";
          note = `전월 ${prev.name} → 당월 ${current.name}`;
        }
      }

      result.set(key, {
        ...canonical,
        currentAmount: current?.amount || 0,
        prevAmount: prev?.amount || 0,
        statusLabel,
        note,
      });
    });

    return Array.from(result.values()).sort((a, b) =>
      a.name.localeCompare(b.name, "ko-KR", { numeric: true }),
    );
  }, [currentRows, prevMonthRows]);

  const prevYearCompareRows = useMemo(() => {
    const existingByCode = new Map(existingRows.map((row) => [norm(row.code), row]));
    const existingByName = new Map(
      existingRows.map((row) => [normalizeStoreNameKey(row.name), row]),
    );

    return prevYearRows
      .map((row) => {
        const byCode = existingByCode.get(norm(row.code));
        const byName = existingByName.get(normalizeStoreNameKey(row.name));
        const manual = codeMappings.find(
          (mapping) =>
            norm(mapping.oldCode) === norm(row.code) &&
            (!mapping.oldName ||
              normalizeStoreNameKey(mapping.oldName) ===
                normalizeStoreNameKey(row.name)),
        );
        let statusLabel = "기존 리스트에 없음";
        let note = "기존거래처 리스트에 없는 전년동월 거래처입니다.";
        let matchedCode = "";
        let matchedName = "";

        if (manual) {
          statusLabel = "수동 매핑";
          note = "저장된 거래처 코드 매핑이 적용됩니다.";
          matchedCode = manual.currentCode;
          matchedName = manual.currentName;
        } else if (byCode && byName) {
          statusLabel = "일치";
          note = "사업자번호와 거래처명이 모두 같습니다.";
          matchedCode = byCode.code;
          matchedName = byCode.name;
        } else if (byName && !byCode) {
          statusLabel = "사업자번호 다름";
          note = `동일 거래처명 기준: ${row.code} ↔ ${byName.code}`;
          matchedCode = byName.code;
          matchedName = byName.name;
        } else if (byCode && !byName) {
          statusLabel = "거래처명 다름";
          note = `동일 사업자번호 기준: ${row.name} ↔ ${byCode.name}`;
          matchedCode = byCode.code;
          matchedName = byCode.name;
        }

        return { ...row, statusLabel, note, matchedCode, matchedName };
      })
      .sort((a, b) =>
        a.statusLabel.localeCompare(b.statusLabel, "ko-KR") ||
        a.name.localeCompare(b.name, "ko-KR", { numeric: true }),
      );
  }, [prevYearRows, existingRows, codeMappings]);

  const totalRows = useMemo(() => {
    const map = new Map<string, Store>();
    existingRows.forEach((row) => {
      const saved = stores.find((store) => store.code === row.code);
      map.set(row.code, {
        code: row.code,
        name: row.name,
        channel: saved?.channel || row.channel || "미지정",
        manager: saved?.manager || row.manager || "",
        storeType: saved?.storeType || row.storeType || "비매장",
        brand: displayBrand(saved?.brand || row.brand),
        status: "거래중",
      });
    });

    prevYearCompareRows.forEach((row) => {
      const targetCode = row.matchedCode || row.code;
      const targetName = row.matchedName || row.name;
      if (map.has(targetCode)) return;
      const saved = stores.find((store) => store.code === targetCode);
      map.set(targetCode, {
        code: targetCode,
        name: targetName,
        channel: saved?.channel || row.channel || "미지정",
        manager: saved?.manager || row.manager || "",
        storeType: saved?.storeType || row.storeType || "비매장",
        brand: displayBrand(saved?.brand || row.brand),
        status: "거래종료",
      });
    });

    return Array.from(map.values()).sort((a, b) =>
      a.name.localeCompare(b.name, "ko-KR", { numeric: true }),
    );
  }, [existingRows, prevYearCompareRows, stores]);

  const visibleExistingRows = existingRows.filter((row) =>
    `${row.code} ${row.name} ${row.statusLabel} ${row.note}`
      .toLowerCase()
      .includes(normalizedSearch),
  );
  const visiblePrevYearRows = prevYearCompareRows.filter((row) =>
    `${row.code} ${row.name} ${row.statusLabel} ${row.note}`
      .toLowerCase()
      .includes(normalizedSearch),
  );
  const visibleTotalRows = totalRows.filter((row) =>
    `${row.code} ${row.name} ${row.channel} ${row.manager} ${row.brand}`
      .toLowerCase()
      .includes(normalizedSearch),
  );
  const newStoreCodes = new Set(
    existingRows
      .filter((row) => row.statusLabel === "당월 신규")
      .map((row) => row.code),
  );
  const needsReferenceSetup = (row: Store) => {
    if (otherTab === "브랜드 관리") {
      const brand = displayBrand(row.brand).trim();
      return newStoreCodes.has(row.code) || !brand || brand === "당월 신규 거래처" || brand === "미지정";
    }
    if (otherTab === "채널 관리") {
      const channel = row.channel.trim();
      return newStoreCodes.has(row.code) || !channel || channel === "미지정";
    }
    return newStoreCodes.has(row.code) || !row.manager.trim();
  };
  const visibleOtherRows = showNewOnly && otherTab !== "담당자 관리"
    ? visibleTotalRows.filter(needsReferenceSetup)
    : visibleTotalRows;
  const normalizeManager = (value: string) => value.trim().toUpperCase();
  const deletedManagerSet = new Set(deletedManagers.map(normalizeManager));
  const managerOptions = Array.from(
    new Set(
      [...savedManagers, ...MANAGERS, ...totalRows.map((row) => row.manager)]
        .map(normalizeManager)
        .filter((value) => Boolean(value) && !deletedManagerSet.has(value)),
    ),
  ).sort((a, b) => a.localeCompare(b, "ko-KR"));
  const channel1Options = Array.from(
    new Map(
      [...savedChannel1Options, ...totalRows.map((row) => row.channel)]
        .map((value) => value.trim())
        .filter(Boolean)
        .map((value) => [value.toLowerCase(), value] as const),
    ).values(),
  ).sort((a, b) => a.localeCompare(b, "ko-KR"));
  const brandOptions = Array.from(
    new Set(totalRows.map((row) => displayBrand(row.brand)).filter(Boolean)),
  ).sort((a, b) => a.localeCompare(b, "ko-KR", { numeric: true }));
  const allVisibleSelected =
    visibleOtherRows.length > 0 &&
    visibleOtherRows.every((row) => selectedCodes.has(row.code));

  function toggleSelected(code: string) {
    setSelectedCodes((previous) => {
      const next = new Set(previous);
      if (next.has(code)) next.delete(code);
      else next.add(code);
      return next;
    });
  }

  function toggleVisibleSelection() {
    setSelectedCodes((previous) => {
      const next = new Set(previous);
      if (allVisibleSelected) {
        visibleOtherRows.forEach((row) => next.delete(row.code));
      } else {
        visibleOtherRows.forEach((row) => next.add(row.code));
      }
      return next;
    });
  }

  function applyBulkChange(field: "manager" | "brand" | "channel" | "storeType") {
    const rawValue = field === "manager"
      ? bulkManager
      : field === "brand"
        ? bulkBrand
        : field === "channel"
          ? bulkChannel1
          : bulkChannel2;
    const value = field === "manager" ? normalizeManager(rawValue) : rawValue.trim();
    if (!selectedCodes.size) return alert("수정할 거래처를 먼저 선택해주세요.");
    if (!value) return alert("변경할 값을 선택하거나 입력해주세요.");
    const label = field === "manager" ? "담당자" : field === "brand" ? "브랜드" : field === "channel" ? "채널 1" : "채널 2";
    if (!window.confirm(`선택한 ${selectedCodes.size.toLocaleString("ko-KR")}개 거래처의 ${label}를 '${value}'(으)로 일괄 변경할까요?`)) return;
    const nextStores = totalRows.map((row) => selectedCodes.has(row.code) ? { ...row, [field]: value } : row);
    setStores(nextStores);
    if (field === "manager" && !savedManagers.some((item) => normalizeManager(item) === value)) {
      setSavedManagers([...savedManagers, value]);
    }
    if (field === "channel" && !savedChannel1Options.some((item) => item.toLowerCase() === value.toLowerCase())) {
      setSavedChannel1Options([...savedChannel1Options, value]);
    }
    setSelectedCodes(new Set());
    alert(`${label}를 ${value}(으)로 일괄 변경했습니다.`);
  }

  function addManager() {
    const value = normalizeManager(newManager);
    if (!value) return alert("추가할 담당자명을 입력해주세요.");
    if (managerOptions.some((item) => normalizeManager(item) === value)) {
      setBulkManager(value);
      setNewManager("");
      return alert("같은 철자의 담당자가 이미 있어 기존 담당자로 선택했습니다.");
    }
    setSavedManagers([...savedManagers, value]);
    setDeletedManagers(deletedManagers.filter((item) => normalizeManager(item) !== value));
    setBulkManager(value);
    setNewManager("");
  }

  function deleteManager() {
    const value = normalizeManager(bulkManager);
    if (!value) return alert("삭제할 담당자를 먼저 선택해주세요.");
    const assignedCount = totalRows.filter((row) => normalizeManager(row.manager) === value).length;
    const assignedMessage = assignedCount
      ? `\n현재 ${assignedCount.toLocaleString("ko-KR")}개 거래처에 지정되어 있습니다. 삭제하면 해당 거래처의 담당자는 미지정으로 변경됩니다.`
      : "";
    if (!window.confirm(`담당자 '${value}'를 삭제할까요?${assignedMessage}`)) return;
    setSavedManagers(savedManagers.filter((item) => normalizeManager(item) !== value));
    if (!deletedManagerSet.has(value)) setDeletedManagers([...deletedManagers, value]);
    if (assignedCount) {
      setStores(totalRows.map((row) => normalizeManager(row.manager) === value ? { ...row, manager: "" } : row));
    }
    setBulkManager("");
    setSelectedCodes(new Set());
  }

  function addChannel1Option() {
    const value = newChannel1.trim();
    if (!value) return alert("추가할 업종명을 입력해주세요.");
    const existing = channel1Options.find((item) => item.toLowerCase() === value.toLowerCase());
    if (existing) {
      setBulkChannel1(existing);
      setNewChannel1("");
      return alert("같은 업종이 이미 있어 기존 업종으로 선택했습니다.");
    }
    setSavedChannel1Options([...savedChannel1Options, value]);
    setBulkChannel1(value);
    setNewChannel1("");
  }

  function saveTotalList() {
    if (!window.confirm(
      `총 거래처 리스트 ${totalRows.length.toLocaleString("ko-KR")}건을 최종 거래처 기준으로 저장할까요?\n저장 후 매출현황과 거래처별 상세가 이 리스트를 기준으로 표시됩니다.`,
    )) return;
    setStores(totalRows);
    alert(`총 거래처 리스트 ${totalRows.length.toLocaleString("ko-KR")}건을 저장했습니다.`);
  }

  const badgeClass = (label: string) => {
    if (["동일", "일치", "수동 매핑"].includes(label))
      return "bg-emerald-100 text-emerald-800";
    if (["당월 신규", "기존 리스트에 없음"].includes(label))
      return "bg-blue-100 text-blue-800";
    if (["당월 미거래"].includes(label))
      return "bg-slate-200 text-slate-700";
    return "bg-amber-100 text-amber-900";
  };

  return (
    <div className="grid gap-3 lg:grid-cols-[150px_minmax(0,1fr)]">
      <aside className="h-fit rounded-xl border border-slate-300 bg-white p-2 shadow-sm">
        <div className="px-2 pb-2 text-[11px] font-extrabold uppercase tracking-wide text-slate-400">
          {section === "list" ? "거래처 리스트" : "기준정보"}
        </div>
        <div className="space-y-1">
          {(section === "list"
            ? (["기존거래처 리스트", "전년동월 리스트", "총 거래처 리스트"] as const)
            : (["담당자 관리", "브랜드 관리", "채널 관리"] as const)
          ).map((menu) => {
            const active = section === "list" ? listTab === menu : otherTab === menu;
            return (
              <button
                key={menu}
                type="button"
                onClick={() => {
                  if (section === "list") setListTab(menu as ListTab);
                  else {
                    setListTab("기타 관리");
                    setOtherTab(menu as typeof otherTab);
                    setSelectedCodes(new Set());
                    setShowNewOnly(false);
                  }
                }}
                className={`w-full rounded-lg border-l-4 px-3 py-2 text-left text-xs font-bold transition ${active ? "border-blue-600 bg-blue-50 text-blue-800" : "border-transparent text-slate-600 hover:bg-slate-50 hover:text-slate-900"}`}
              >
                {menu.replace(" 리스트", "").replace(" 관리", "")}
              </button>
            );
          })}
        </div>
      </aside>

      <main className="min-w-0 space-y-3">
        <div className="rounded-xl border border-slate-300 bg-white p-3 shadow-sm">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <h2 className="text-sm font-extrabold text-slate-900">{section === "list" ? listTab : otherTab}</h2>
            <div className="flex flex-wrap items-center gap-2">
              <input value={search} onChange={(event) => setSearch(event.target.value)} placeholder={section === "reference" ? "거래처명/담당자/브랜드/채널 검색" : "사업자번호/거래처명/상태 검색"} className="h-8 w-[270px] rounded-lg border border-slate-300 px-3 text-xs outline-none focus:border-blue-500" />
              {section === "list" && listTab === "총 거래처 리스트" && (
                <button type="button" onClick={saveTotalList} className="h-8 rounded-lg bg-blue-600 px-3 text-xs font-bold text-white hover:bg-blue-700">총 리스트 저장</button>
              )}
            </div>
          </div>

          {section === "reference" && (
            <>
              <div className="mt-3 flex flex-wrap items-center gap-2 rounded-lg bg-slate-50 p-2">
                <span className="mr-1 text-[11px] font-semibold text-slate-600">선택 {selectedCodes.size.toLocaleString("ko-KR")}건 / 검색결과 {visibleOtherRows.length.toLocaleString("ko-KR")}건</span>
                <button type="button" onClick={toggleVisibleSelection} className="h-8 rounded-lg border border-slate-300 bg-white px-3 text-xs font-bold text-slate-700 hover:bg-slate-100">{allVisibleSelected ? "검색결과 선택 해제" : "검색결과 전체 선택"}</button>
                <button type="button" onClick={() => setSelectedCodes(new Set())} className="h-8 rounded-lg border border-slate-300 bg-white px-3 text-xs font-bold text-slate-700 hover:bg-slate-100">선택 초기화</button>
                {otherTab !== "담당자 관리" && (
                  <button
                    type="button"
                    onClick={() => {
                      setShowNewOnly((previous) => !previous);
                      setSelectedCodes(new Set());
                    }}
                    className={`h-8 rounded-lg border px-3 text-xs font-bold ${showNewOnly ? "border-orange-400 bg-orange-50 text-orange-800" : "border-slate-300 bg-white text-slate-700 hover:bg-slate-100"}`}
                  >
                    {showNewOnly ? "전체 거래처 보기" : "신규·미설정 거래처 보기"}
                  </button>
                )}
                {otherTab === "담당자 관리" && (
                  <>
                    <select value={bulkManager} onChange={(event) => setBulkManager(event.target.value)} className="h-8 min-w-[150px] rounded-lg border border-slate-300 bg-white px-2 text-xs font-semibold"><option value="">변경 담당자 선택</option>{managerOptions.map((manager) => <option key={manager} value={manager}>{manager}</option>)}</select>
                    <button type="button" onClick={() => applyBulkChange("manager")} className="h-8 rounded-lg bg-blue-600 px-3 text-xs font-bold text-white">담당자 일괄 수정</button>
                    <button type="button" onClick={deleteManager} className="h-8 rounded-lg border border-red-300 bg-red-50 px-3 text-xs font-bold text-red-700 hover:bg-red-100">담당자 삭제</button>
                    <span className="mx-1 h-5 w-px bg-slate-300" />
                    <input value={newManager} onChange={(event) => setNewManager(event.target.value)} onKeyDown={(event) => { if (event.key === "Enter") addManager(); }} placeholder="신규 담당자" className="h-8 w-[130px] rounded-lg border border-slate-300 bg-white px-2 text-xs" />
                    <button type="button" onClick={addManager} className="h-8 rounded-lg border border-blue-300 bg-blue-50 px-3 text-xs font-bold text-blue-700">담당자 추가</button>
                  </>
                )}
                {otherTab === "브랜드 관리" && (
                  <>
                    <input list="store-brand-options" value={bulkBrand} onChange={(event) => setBulkBrand(event.target.value)} placeholder="변경 브랜드 입력 또는 선택" className="h-8 min-w-[220px] rounded-lg border border-slate-300 bg-white px-2 text-xs font-semibold" />
                    <datalist id="store-brand-options">{brandOptions.map((brand) => <option key={brand} value={brand} />)}</datalist>
                    <button type="button" onClick={() => applyBulkChange("brand")} className="h-8 rounded-lg bg-blue-600 px-3 text-xs font-bold text-white">브랜드 일괄 수정</button>
                  </>
                )}
                {otherTab === "채널 관리" && (
                  <>
                    <div className="flex rounded-lg border border-slate-300 bg-white p-0.5">{(["채널 1", "채널 2"] as const).map((item) => <button key={item} type="button" onClick={() => setChannelTab(item)} className={`rounded-md px-3 py-1.5 text-xs font-bold ${channelTab === item ? "bg-slate-800 text-white" : "text-slate-600"}`}>{item}</button>)}</div>
                    {channelTab === "채널 1" ? (
                      <>
                        <select value={bulkChannel1} onChange={(event) => setBulkChannel1(event.target.value)} className="h-8 min-w-[150px] rounded-lg border border-slate-300 bg-white px-2 text-xs font-semibold"><option value="">업종 선택</option>{channel1Options.map((channel) => <option key={channel} value={channel}>{channel}</option>)}</select>
                        <button type="button" onClick={() => applyBulkChange("channel")} className="h-8 rounded-lg bg-blue-600 px-3 text-xs font-bold text-white">채널 1 일괄 수정</button>
                        <span className="mx-1 h-5 w-px bg-slate-300" />
                        <input value={newChannel1} onChange={(event) => setNewChannel1(event.target.value)} onKeyDown={(event) => { if (event.key === "Enter") addChannel1Option(); }} placeholder="신규 업종" className="h-8 w-[130px] rounded-lg border border-slate-300 bg-white px-2 text-xs" />
                        <button type="button" onClick={addChannel1Option} className="h-8 rounded-lg border border-blue-300 bg-blue-50 px-3 text-xs font-bold text-blue-700">업종 추가</button>
                      </>
                    ) : (
                      <>
                        <select value={bulkChannel2} onChange={(event) => setBulkChannel2(event.target.value as "" | "매장" | "비매장")} className="h-8 min-w-[140px] rounded-lg border border-slate-300 bg-white px-2 text-xs font-semibold"><option value="">채널 2 선택</option><option value="매장">매장</option><option value="비매장">비매장</option></select>
                        <button type="button" onClick={() => applyBulkChange("storeType")} className="h-8 rounded-lg bg-blue-600 px-3 text-xs font-bold text-white">채널 2 일괄 수정</button>
                      </>
                    )}
                  </>
                )}
              </div>
            </>
          )}
        </div>

      <div className="min-h-0 overflow-hidden rounded-2xl border border-slate-300 bg-white shadow-sm">
        <div className="max-h-[65vh] overflow-auto">
          {listTab === "기존거래처 리스트" && (
            <table className="w-full min-w-[1250px] border-separate border-spacing-0 text-center text-xs whitespace-nowrap">
              <thead>
                <tr className="bg-slate-100">
                  <th className="sticky top-0 z-20 border border-slate-300 bg-slate-100 px-3 py-2">사업자번호</th>
                  <th className="sticky top-0 z-20 border border-slate-300 bg-slate-100 px-3 py-2">거래처명</th>
                  <th className="sticky top-0 z-20 border border-slate-300 bg-blue-50 px-3 py-2">전월 매출</th>
                  <th className="sticky top-0 z-20 border border-slate-300 bg-orange-50 px-3 py-2">당월 매출</th>
                  <th className="sticky top-0 z-20 border border-slate-300 bg-slate-100 px-3 py-2">비교 결과</th>
                  <th className="sticky top-0 z-20 border border-slate-300 bg-slate-100 px-3 py-2">확인 내용</th>
                </tr>
              </thead>
              <tbody>
                {visibleExistingRows.map((row) => (
                  <tr key={`${row.code}|${row.name}`} className="hover:bg-orange-50/50">
                    <td className="border border-slate-300 px-3 py-2">{row.code}</td>
                    <td className="border border-slate-300 px-3 py-2 text-left font-semibold">{row.name}</td>
                    <td className="border border-slate-300 px-3 py-2 text-right font-semibold">{won(row.prevAmount)}</td>
                    <td className="border border-slate-300 px-3 py-2 text-right font-semibold">{won(row.currentAmount)}</td>
                    <td className="border border-slate-300 px-3 py-2"><span className={`rounded-full px-3 py-1 font-bold ${badgeClass(row.statusLabel)}`}>{row.statusLabel}</span></td>
                    <td className="border border-slate-300 px-3 py-2 text-left text-slate-600">{row.note}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}

          {listTab === "전년동월 리스트" && (
            <table className="w-full min-w-[1350px] border-separate border-spacing-0 text-center text-xs whitespace-nowrap">
              <thead>
                <tr className="bg-slate-100">
                  <th className="sticky top-0 z-20 border border-slate-300 bg-slate-100 px-3 py-2">전년 사업자번호</th>
                  <th className="sticky top-0 z-20 border border-slate-300 bg-slate-100 px-3 py-2">전년 거래처명</th>
                  <th className="sticky top-0 z-20 border border-slate-300 bg-emerald-50 px-3 py-2">전년동월 매출</th>
                  <th className="sticky top-0 z-20 border border-slate-300 bg-slate-100 px-3 py-2">비교 결과</th>
                  <th className="sticky top-0 z-20 border border-slate-300 bg-slate-100 px-3 py-2">기존 리스트 연결</th>
                  <th className="sticky top-0 z-20 border border-slate-300 bg-slate-100 px-3 py-2">확인 내용</th>
                </tr>
              </thead>
              <tbody>
                {visiblePrevYearRows.map((row) => (
                  <tr key={`${row.code}|${row.name}`} className="hover:bg-orange-50/50">
                    <td className="border border-slate-300 px-3 py-2">{row.code}</td>
                    <td className="border border-slate-300 px-3 py-2 text-left font-semibold">{row.name}</td>
                    <td className="border border-slate-300 px-3 py-2 text-right font-semibold">{won(row.amount)}</td>
                    <td className="border border-slate-300 px-3 py-2"><span className={`rounded-full px-3 py-1 font-bold ${badgeClass(row.statusLabel)}`}>{row.statusLabel}</span></td>
                    <td className="border border-slate-300 px-3 py-2 text-left">{row.matchedCode ? `${row.matchedCode} / ${row.matchedName}` : "-"}</td>
                    <td className="border border-slate-300 px-3 py-2 text-left text-slate-600">{row.note}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}

          {listTab === "총 거래처 리스트" && (
            <table className="w-full min-w-[1400px] border-separate border-spacing-0 text-center text-xs whitespace-nowrap">
              <thead>
                <tr className="bg-slate-100">
                  <th className="sticky top-0 z-20 border border-slate-300 bg-slate-100 px-3 py-2">사업자번호</th>
                  <th className="sticky top-0 z-20 border border-slate-300 bg-slate-100 px-3 py-2">거래처명</th>
                  <th className="sticky top-0 z-20 border border-slate-300 bg-slate-100 px-3 py-2">담당자</th>
                  <th className="sticky top-0 z-20 border border-slate-300 bg-slate-100 px-3 py-2">채널</th>
                  <th className="sticky top-0 z-20 border border-slate-300 bg-slate-100 px-3 py-2">매장/비매장</th>
                  <th className="sticky top-0 z-20 border border-slate-300 bg-slate-100 px-3 py-2">브랜드</th>
                  <th className="sticky top-0 z-20 border border-slate-300 bg-slate-100 px-3 py-2">상태</th>
                </tr>
              </thead>
              <tbody>
                {visibleTotalRows.map((row) => (
                  <tr key={row.code} className="hover:bg-orange-50/50">
                    <td className="border border-slate-300 px-3 py-2">{row.code}</td>
                    <td className="border border-slate-300 px-3 py-2 text-left font-semibold">{row.name}</td>
                    <td className="border border-slate-300 px-3 py-2 font-bold">{row.manager ? normalizeManager(row.manager) : "미지정"}</td>
                    <td className="border border-slate-300 px-3 py-2">{row.channel}</td>
                    <td className="border border-slate-300 px-3 py-2">{row.storeType}</td>
                    <td className="border border-slate-300 px-3 py-2 text-left">{row.brand}</td>
                    <td className="border border-slate-300 px-3 py-2"><span className={`rounded-full px-3 py-1 font-bold ${row.status === "거래중" ? "bg-emerald-100 text-emerald-800" : "bg-slate-200 text-slate-700"}`}>{row.status}</span></td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}

          {section === "reference" && listTab === "기타 관리" && (
            <table className="w-full min-w-[1150px] border-separate border-spacing-0 text-center text-xs whitespace-nowrap">
              <thead>
                <tr className="bg-slate-100">
                  <th className="sticky top-0 z-20 w-12 border border-slate-300 bg-slate-100 px-2 py-2">
                    <input
                      type="checkbox"
                      checked={allVisibleSelected}
                      onChange={toggleVisibleSelection}
                      aria-label="검색결과 전체 선택"
                      className="h-4 w-4 accent-blue-600"
                    />
                  </th>
                  <th className="sticky top-0 z-20 border border-slate-300 bg-slate-100 px-3 py-2">사업자번호</th>
                  <th className="sticky top-0 z-20 border border-slate-300 bg-slate-100 px-3 py-2">거래처명</th>
                  <th className="sticky top-0 z-20 border border-slate-300 bg-blue-50 px-3 py-2">현재 담당자</th>
                  <th className="sticky top-0 z-20 border border-slate-300 bg-orange-50 px-3 py-2">현재 브랜드</th>
                  <th className="sticky top-0 z-20 border border-slate-300 bg-slate-100 px-3 py-2">채널</th>
                  <th className="sticky top-0 z-20 border border-slate-300 bg-slate-100 px-3 py-2">매장/비매장</th>
                  <th className="sticky top-0 z-20 border border-slate-300 bg-slate-100 px-3 py-2">상태</th>
                </tr>
              </thead>
              <tbody>
                {visibleOtherRows.map((row) => {
                  const checked = selectedCodes.has(row.code);
                  return (
                    <tr
                      key={row.code}
                      className={checked ? "bg-blue-50" : "hover:bg-orange-50/50"}
                      onClick={() => toggleSelected(row.code)}
                    >
                      <td className="border border-slate-300 px-2 py-2">
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={() => toggleSelected(row.code)}
                          onClick={(event) => event.stopPropagation()}
                          aria-label={`${row.name} 선택`}
                          className="h-4 w-4 accent-blue-600"
                        />
                      </td>
                      <td className="border border-slate-300 px-3 py-2">{row.code}</td>
                      <td className="border border-slate-300 px-3 py-2 text-left font-semibold">{row.name}</td>
                      <td className="border border-slate-300 px-3 py-2 font-bold">{row.manager ? normalizeManager(row.manager) : "미지정"}</td>
                      <td className="border border-slate-300 px-3 py-2 text-left">{row.brand || "미지정"}</td>
                      <td className="border border-slate-300 px-3 py-2">{row.channel}</td>
                      <td className="border border-slate-300 px-3 py-2">{row.storeType}</td>
                      <td className="border border-slate-300 px-3 py-2">
                        <span className={`rounded-full px-3 py-1 font-bold ${row.status === "거래중" ? "bg-emerald-100 text-emerald-800" : "bg-slate-200 text-slate-700"}`}>
                          {row.status}
                        </span>
                      </td>
                    </tr>
                  );
                })}
                {!visibleOtherRows.length && (
                  <tr>
                    <td colSpan={8} className="border border-slate-300 px-4 py-12 text-center text-sm text-slate-500">
                      검색 조건에 맞는 거래처가 없습니다.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          )}
        </div>
      </div>
      </main>
    </div>
  );
}

function MappingPage({
  stores,
  setStores,
  sales,
  month,
  codeMappings,
  setCodeMappings,
}: {
  stores: Store[];
  setStores: (v: Store[]) => void;
  sales: SalesRecord[];
  month: string;
  codeMappings: StoreCodeMapping[];
  setCodeMappings: (v: StoreCodeMapping[]) => void;
}) {
  const empty: Store = {
    code: "",
    name: "",
    channel: "도매",
    manager: "",
    storeType: "비매장",
    brand: "당월 신규 거래처",
    status: "거래중",
  };
  const [form, setForm] = useState<Store>(empty);
  const [search, setSearch] = useState("");
  const [searchDraft, setSearchDraft] = useState("");
  const [statusFilter, setStatusFilter] = useState<
    "all" | "active" | "inactive"
  >("active");
  const [mappingListOpen, setMappingListOpen] = useState(false);
  const [mappingModalOpen, setMappingModalOpen] = useState(false);
  const rows = stores
    .filter(
      (s) =>
        statusFilter === "all" ||
        (statusFilter === "active" && s.status === "거래중") ||
        (statusFilter === "inactive" && s.status === "거래종료"),
    )
    .filter((s) =>
      `${s.code} ${s.name} ${s.channel} ${s.manager} ${s.brand}`
        .toLowerCase()
        .includes(search.toLowerCase()),
    );
  const activeCount = stores.filter((s) => s.status === "거래중").length;
  const inactiveCount = stores.filter((s) => s.status === "거래종료").length;

  function save() {
    if (!form.code || !form.name)
      return alert("거래처코드와 거래처명은 필수입니다.");
    const exists = stores.some((s) => s.code === form.code);
    setStores(
      exists
        ? stores.map((s) => (s.code === form.code ? form : s))
        : [...stores, form],
    );
    setForm(empty);
  }

  function toggleStatus(store: Store, e: React.MouseEvent<HTMLButtonElement>) {
    e.stopPropagation();
    const nextStatus: Store["status"] =
      store.status === "거래중" ? "거래종료" : "거래중";
    setStores(
      stores.map((s) =>
        s.code === store.code ? { ...s, status: nextStatus } : s,
      ),
    );
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
    const parsed: Store[] = rows
      .map((r) => {
        const channel = normalizeChannel(r["채널1"] ?? r["채널"]);
        return {
          code: norm(r["거래처코드"] ?? r["거래처 코드"] ?? r["매장코드"]),
          name: norm(r["거래처명"] ?? r["매장명"]),
          channel,
          manager: norm(r["담당자"]) as Manager,
          storeType: normalizeStoreType(r["채널2"] ?? r["매장구분"], channel),
          brand: displayBrand(r["브랜드"]),
          status: normalizeStatus(r["거래상태"]),
        };
      })
      .filter((s) => s.code && s.name);

    const map = new Map(stores.map((s) => [s.code, s]));
    parsed.forEach((s) => map.set(s.code, { ...map.get(s.code), ...s }));
    setStores(Array.from(map.values()));
    alert(`거래처 매핑 ${parsed.length}건을 반영했습니다.`);
  }

  const currentSalesStores = useMemo(() => {
    const map = new Map<string, { code: string; name: string }>();
    sales
      .filter(
        (r) =>
          r.period === "current" &&
          inRange(r.saleDate, monthStart(month), monthEnd(month)),
      )
      .forEach((r) => {
        if (!map.has(r.storeCode))
          map.set(r.storeCode, { code: r.storeCode, name: r.storeName });
      });
    return Array.from(map.values()).sort((a, b) =>
      a.name.localeCompare(b.name, "ko-KR"),
    );
  }, [sales, month]);

  const mappingCheckRows = useMemo(() => {
    const currentByCode = new Map(currentSalesStores.map((s) => [s.code, s]));
    const currentByName = new Map<string, { code: string; name: string }>();
    currentSalesStores.forEach((s) => {
      const key = normalizeStoreNameKey(s.name);
      if (key && !currentByName.has(key)) currentByName.set(key, s);
    });

    const sourceMap = new Map<
      string,
      { period: PeriodType; code: string; name: string; amount: number }
    >();
    sales
      .filter((r) => r.period === "prevYear" && r.refMonth === month)
      .forEach((r) => {
        const key = `${r.period}|${r.storeCode}|${r.storeName}`;
        const item = sourceMap.get(key) || {
          period: r.period,
          code: r.storeCode,
          name: r.storeName,
          amount: 0,
        };
        item.amount += r.salesAmount;
        sourceMap.set(key, item);
      });

    return Array.from(sourceMap.values())
      .map((r) => {
        const byCode = currentByCode.get(r.code);
        const byName = currentByName.get(normalizeStoreNameKey(r.name));
        const manual = codeMappings.find(
          (m) =>
            norm(m.oldCode) === norm(r.code) &&
            (!m.oldName ||
              normalizeStoreNameKey(m.oldName) ===
                normalizeStoreNameKey(r.name)),
        );
        let category = "수동 매핑 필요";
        let reason = "당월 매출이 없어 전년동월 업로드 거래처 기준으로 표시";
        let targetCode = "";
        let targetName = "";

        if (manual) {
          category = "수동 매핑 완료";
          reason = "사용자가 현재 거래처로 직접 매핑함";
          targetCode = manual.currentCode;
          targetName = manual.currentName;
        } else if (
          byCode &&
          normalizeStoreNameKey(byCode.name) === normalizeStoreNameKey(r.name)
        ) {
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
      })
      .sort(
        (a, b) =>
          a.category.localeCompare(b.category, "ko-KR") ||
          a.name.localeCompare(b.name, "ko-KR"),
      );
  }, [sales, month, currentSalesStores, codeMappings]);

  const mappingSummary = useMemo(() => {
    return mappingCheckRows.reduce(
      (acc, row) => {
        acc[row.category] = (acc[row.category] || 0) + 1;
        return acc;
      },
      {} as Record<string, number>,
    );
  }, [mappingCheckRows]);

  function saveManualMapping(row: {
    code: string;
    name: string;
    targetCode: string;
    targetName: string;
  }) {
    const currentCode = window.prompt(
      "현재 거래처코드를 입력하세요.",
      row.targetCode || "",
    );
    if (!currentCode) return;
    const matched =
      currentSalesStores.find((s) => s.code === currentCode) ||
      stores.find((s) => s.code === currentCode);
    const currentName = window.prompt(
      "현재 거래처명을 입력하세요.",
      row.targetName || matched?.name || "",
    );
    if (!currentName) return;
    const next: StoreCodeMapping = {
      id: `${row.code}|${row.name}|${currentCode}`,
      oldCode: row.code,
      oldName: row.name,
      currentCode,
      currentName,
    };
    setCodeMappings([
      ...codeMappings.filter(
        (m) =>
          !(
            norm(m.oldCode) === norm(row.code) &&
            normalizeStoreNameKey(m.oldName) === normalizeStoreNameKey(row.name)
          ),
      ),
      next,
    ]);
  }

  function deleteManualMapping(row: { code: string; name: string }) {
    setCodeMappings(
      codeMappings.filter(
        (m) =>
          !(
            norm(m.oldCode) === norm(row.code) &&
            normalizeStoreNameKey(m.oldName) === normalizeStoreNameKey(row.name)
          ),
      ),
    );
  }

  return (
    <div className="rounded-xl border border-gray-300/70 bg-white/80 p-4 shadow-sm backdrop-blur">
      <div className="mb-3 flex flex-col gap-2 lg:flex-row lg:items-center lg:justify-between">
        <div>
          <h2 className="text-base font-bold">거래처 매핑관리</h2>
          <p className="mt-1 text-xs text-slate-500">
            채널/담당자/브랜드 정보를 거래처코드 기준으로 관리합니다. 전년동월
            매출은 당월 매출에 같은 거래처코드나 거래처명이 있으면 당월 거래처
            기준으로 합산하고, 당월 매출이 없으면 전년동월 업로드 거래처
            기준으로 별도 표시합니다.
          </p>
        </div>
        <label className="shrink-0 cursor-pointer rounded-md bg-green-600 px-2 py-1 text-[11px] font-semibold text-white">
          매핑 엑셀 업로드
          <input
            type="file"
            accept=".xlsx,.xls,.csv"
            className="hidden"
            onChange={(e) => upload(e.target.files?.[0] || null)}
          />
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
            <button
              type="button"
              onClick={() => {
                setStatusFilter("all");
                setMappingListOpen(true);
              }}
              className={`px-3 py-1.5 ${statusFilter === "all" ? "bg-slate-900 text-white" : "text-slate-600 hover:bg-slate-50"}`}
            >
              전체 {stores.length.toLocaleString("ko-KR")}
            </button>
            <button
              type="button"
              onClick={() => {
                setStatusFilter("active");
                setMappingListOpen(true);
              }}
              className={`border-l border-slate-200 px-3 py-1.5 ${statusFilter === "active" ? "bg-emerald-600 text-white" : "text-slate-600 hover:bg-slate-50"}`}
            >
              활성 {activeCount.toLocaleString("ko-KR")}
            </button>
            <button
              type="button"
              onClick={() => {
                setStatusFilter("inactive");
                setMappingListOpen(true);
              }}
              className={`border-l border-slate-200 px-3 py-1.5 ${statusFilter === "inactive" ? "bg-slate-600 text-white" : "text-slate-600 hover:bg-slate-50"}`}
            >
              비활성 {inactiveCount.toLocaleString("ko-KR")}
            </button>
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
          <span className="text-xs font-semibold text-slate-500">
            표시 {rows.length.toLocaleString("ko-KR")}건
          </span>
        </div>
      </div>

      {mappingListOpen && (
        <div className="mb-3 rounded-xl border border-gray-300 bg-slate-50 p-3">
          <div className="mb-2 text-xs font-bold text-slate-700">
            신규 매장 추가 / 거래처 수정
          </div>
          <div className="flex flex-wrap items-center gap-1.5">
            <input
              value={form.brand}
              onChange={(e) => setForm({ ...form, brand: e.target.value })}
              placeholder="브랜드"
              className="h-8 w-[150px] rounded-md border bg-white px-2 py-1 text-xs"
            />
            <input
              value={form.code}
              onChange={(e) => setForm({ ...form, code: e.target.value })}
              placeholder="거래처코드"
              className="h-8 w-[140px] rounded-md border bg-white px-2 py-1 text-xs"
            />
            <input
              value={form.name}
              onChange={(e) => setForm({ ...form, name: e.target.value })}
              placeholder="거래처명"
              className="h-8 w-[220px] rounded-md border bg-white px-2 py-1 text-xs"
            />
            <select
              value={form.channel}
              onChange={(e) =>
                setForm({ ...form, channel: e.target.value as Channel })
              }
              className="h-8 w-[110px] rounded-md border bg-white px-2 py-1 text-xs"
            >
              <option value="">채널1</option>
              {CHANNELS.map((c) => (
                <option key={c}>{c}</option>
              ))}
            </select>
            <select
              value={form.storeType}
              onChange={(e) =>
                setForm({ ...form, storeType: e.target.value as StoreType })
              }
              className="h-8 w-[105px] rounded-md border bg-white px-2 py-1 text-xs"
            >
              <option>매장</option>
              <option>비매장</option>
            </select>
            <select
              value={form.manager}
              onChange={(e) =>
                setForm({ ...form, manager: e.target.value as Manager })
              }
              className="h-8 w-[100px] rounded-md border bg-white px-2 py-1 text-xs"
            >
              <option value="">담당자</option>
              {MANAGERS.map((m) => (
                <option key={m}>{m}</option>
              ))}
            </select>
            <button
              onClick={save}
              className="h-8 w-[70px] rounded-md bg-blue-600 px-2 py-1 text-xs font-semibold text-white"
            >
              저장
            </button>
            <button
              type="button"
              onClick={() => setForm(empty)}
              className="h-8 rounded-md border border-slate-300 bg-white px-3 py-1 text-xs font-semibold text-slate-700"
            >
              초기화
            </button>
            <span className="text-[11px] font-medium text-slate-500">
              신규 추가 시 상태는 자동으로 활성 처리됩니다. 기존 거래처 수정 시
              상태는 유지됩니다.
            </span>
          </div>
        </div>
      )}

      {mappingModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/55 p-4">
          <div className="flex max-h-[92vh] w-full max-w-[96vw] flex-col rounded-2xl bg-white shadow-2xl">
            <div className="flex items-center justify-between border-b border-gray-300 p-4">
              <div>
                <h3 className="text-lg font-bold text-slate-900">
                  전년동월 거래처 매핑 검증
                </h3>
                <p className="mt-1 text-xs text-slate-500">
                  검증표는 필요할 때만 열어서 수동 매핑을 설정합니다.
                </p>
              </div>
              <button
                type="button"
                onClick={() => setMappingModalOpen(false)}
                className="rounded-lg border border-slate-300 bg-white px-3 py-2 text-xs font-semibold text-slate-700 hover:bg-slate-50"
              >
                닫기
              </button>
            </div>
            <div className="min-h-0 flex-1 overflow-auto bg-white pb-10 scroll-pb-10 px-4 pb-4 pt-0">
              <div className="rounded-xl border border-gray-300 bg-slate-50 p-3">
                <div className="mb-2 flex flex-col gap-1 lg:flex-row lg:items-center lg:justify-between">
                  <div>
                    <h3 className="text-sm font-bold text-slate-900">
                      전년동월 거래처 매핑 검증
                    </h3>
                    <p className="mt-1 text-xs text-slate-500">
                      전년동월 업로드 거래처를 당월 매출 거래처 기준으로
                      비교합니다. 코드나 거래처명 중 하나라도 당월 매출과 같으면
                      당월 거래처 기준으로 자동 합산되고, 당월 매출이 없으면
                      전년동월 업로드 거래처 기준으로 표시됩니다.
                    </p>
                  </div>
                  <div className="flex flex-wrap gap-2 text-xs font-semibold text-slate-600">
                    <span>자동 {mappingSummary["자동 매핑"] || 0}건</span>
                    <span>
                      수동필요 {mappingSummary["수동 매핑 필요"] || 0}건
                    </span>
                    <span>
                      수동완료 {mappingSummary["수동 매핑 완료"] || 0}건
                    </span>
                  </div>
                </div>
                <div className="max-h-[72vh] overflow-auto">
                  <table className="w-full min-w-[960px] border-separate border-spacing-0 border border-gray-300 bg-white text-xs">
                    <thead>
                      <tr className="bg-slate-100">
                        <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-1.5 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">
                          구분
                        </th>
                        <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-1.5 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">
                          업로드구분
                        </th>
                        <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-1.5 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">
                          업로드코드
                        </th>
                        <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-1.5 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">
                          업로드거래처명
                        </th>
                        <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-1.5 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">
                          업로드매출금액
                        </th>
                        <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-1.5 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">
                          당월코드
                        </th>
                        <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-1.5 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">
                          당월거래처명
                        </th>
                        <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-1.5 text-left font-bold shadow-[0_2px_0_0_#e2e8f0]">
                          매핑방식
                        </th>
                        <th className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-1.5 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">
                          관리
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {mappingCheckRows.length === 0 ? (
                        <tr>
                          <td
                            colSpan={9}
                            className="border p-5 text-center text-slate-500"
                          >
                            검증할 전년동월 매출 데이터가 없습니다.
                          </td>
                        </tr>
                      ) : (
                        mappingCheckRows.map((r) => (
                          <tr
                            key={`${r.period}|${r.code}|${r.name}`}
                            className="hover:bg-slate-50"
                          >
                            <td className="border px-2 py-1.5 font-semibold">
                              {r.category}
                            </td>
                            <td className="border px-2 py-1.5">전년동월</td>
                            <td className="border px-2 py-1.5">{r.code}</td>
                            <td className="border px-2 py-1.5 font-semibold">
                              {r.name}
                            </td>
                            <td className="border px-2 py-1.5 text-right font-semibold">
                              {won(r.amount)}
                            </td>
                            <td className="border px-2 py-1.5">
                              {r.targetCode || "-"}
                            </td>
                            <td className="border px-2 py-1.5">
                              {r.targetName || "-"}
                            </td>
                            <td className="border px-2 py-1.5">{r.reason}</td>
                            <td className="sticky top-0 z-[80] border border-slate-300 bg-white px-2 py-1.5 text-right font-bold shadow-[0_2px_0_0_#e2e8f0]">
                              {r.category === "수동 매핑 완료" ? (
                                <button
                                  type="button"
                                  onClick={() => deleteManualMapping(r)}
                                  className="rounded-md border border-red-200 bg-red-50 px-2 py-0.5 text-[11px] font-bold text-red-600 hover:bg-red-100"
                                >
                                  해제
                                </button>
                              ) : r.category === "수동 매핑 필요" ? (
                                <button
                                  type="button"
                                  onClick={() => saveManualMapping(r)}
                                  className="rounded-md bg-blue-600 px-2 py-0.5 text-[11px] font-bold text-white hover:bg-blue-700"
                                >
                                  수동매핑
                                </button>
                              ) : (
                                <span className="text-slate-400">자동</span>
                              )}
                            </td>
                          </tr>
                        ))
                      )}
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
                <th className="border border-gray-300 bg-slate-100 px-2 py-1.5 text-left text-xs font-bold">
                  거래처코드
                </th>
                <th className="border border-gray-300 bg-slate-100 px-2 py-1.5 text-left text-xs font-bold">
                  거래처명
                </th>
                <th className="border border-gray-300 bg-slate-100 px-2 py-1.5 text-left text-xs font-bold">
                  채널
                </th>
                <th className="border border-gray-300 bg-slate-100 px-2 py-1.5 text-left text-xs font-bold">
                  담당자
                </th>
                <th className="border border-gray-300 bg-slate-100 px-2 py-1.5 text-left text-xs font-bold">
                  매장구분
                </th>
                <th className="border border-gray-300 bg-slate-100 px-2 py-1.5 text-left text-xs font-bold">
                  브랜드
                </th>
                <th className="border border-gray-300 bg-slate-100 px-2 py-1.5 text-left text-xs font-bold">
                  상태
                </th>
                <th className="border border-gray-300 bg-slate-100 px-2 py-1.5 text-right text-xs font-bold">
                  관리
                </th>
              </tr>
            </thead>
            <tbody>
              {rows.map((s) => (
                <tr
                  key={s.code}
                  onClick={() => setForm(s)}
                  className="cursor-pointer hover:bg-blue-50"
                >
                  <td
                    className="truncate border border-gray-300 bg-white px-2 py-1.5 text-left align-middle"
                    title={s.code}
                  >
                    {s.code}
                  </td>
                  <td
                    className="truncate border border-gray-300 bg-white px-2 py-1.5 text-left align-middle font-semibold"
                    title={s.name}
                  >
                    {s.name}
                  </td>
                  <td className="border border-gray-300 bg-white px-2 py-1.5 text-left align-middle">
                    {s.channel}
                  </td>
                  <td className="border border-gray-300 bg-white px-2 py-1.5 text-left align-middle">
                    {s.manager || "-"}
                  </td>
                  <td className="border border-gray-300 bg-white px-2 py-1.5 text-left align-middle">
                    {s.storeType}
                  </td>
                  <td
                    className="truncate border border-gray-300 bg-white px-2 py-1.5 text-left align-middle"
                    title={s.brand}
                  >
                    {s.brand}
                  </td>
                  <td className="border border-gray-300 bg-white px-2 py-1.5 text-left align-middle">
                    <button
                      type="button"
                      onClick={(e) => toggleStatus(s, e)}
                      className={`rounded-full px-2 py-0.5 text-[11px] font-semibold ${s.status === "거래중" ? "bg-emerald-50 text-emerald-700 ring-1 ring-emerald-200" : "bg-slate-100 text-slate-500 ring-1 ring-slate-200"}`}
                      title={
                        s.status === "거래중"
                          ? "클릭하면 비활성 처리됩니다."
                          : "클릭하면 다시 활성화됩니다."
                      }
                    >
                      {s.status === "거래중" ? "활성" : "비활성"}
                    </button>
                  </td>
                  <td className="border border-gray-300 bg-white px-2 py-1.5 text-right align-middle">
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

function TargetByTypePage({
  records,
  setRecords,
  month,
}: {
  records: TargetRecord[];
  setRecords: (v: TargetRecord[]) => void;
  month: string;
}) {
  const [targetMonth, setTargetMonth] = useState(month);

  const getAmount = (storeType: StoreType) =>
    records
      .filter(
        (r) =>
          r.month === targetMonth &&
          (r.storeType === storeType ||
            (!r.storeType && storeType === "비매장")),
      )
      .reduce((a, b) => a + b.amount, 0);

  const storeAmount = getAmount("매장");
  const nonStoreAmount = getAmount("비매장");

  function updateAmount(storeType: StoreType, amount: number) {
    const next = { storeType, month: targetMonth, amount };
    setRecords([
      ...records.filter(
        (r) =>
          !(
            r.month === targetMonth &&
            (r.storeType === storeType ||
              (!r.storeType && storeType === "비매장"))
          ),
      ),
      next,
    ]);
  }

  async function upload(file: File | null) {
    if (!file) return;
    const fileRows = await readFileRows(file);
    const parsed = fileRows
      .map((r) => {
        const rawType = norm(
          r["매장구분"] ?? r["구분"] ?? r["타입"] ?? r["storeType"],
        );
        const storeType: StoreType =
          rawType.includes("매장") && !rawType.includes("비매장")
            ? "매장"
            : "비매장";
        return {
          storeType,
          month: monthText(r["기준월"] ?? r["월"] ?? r["년월"] ?? targetMonth),
          amount: num(
            r["Target"] ??
              r["TARGET"] ??
              r["금액"] ??
              r["목표"] ??
              r["목표매출"],
          ),
        };
      })
      .filter((r) => r.month && r.amount > 0);

    setRecords([
      ...records.filter(
        (r) =>
          !parsed.some(
            (p) => p.month === r.month && p.storeType === r.storeType,
          ),
      ),
      ...parsed,
    ]);
    alert(`Target ${parsed.length}건을 반영했습니다.`);
  }

  const excelRows = [
    { 기준월: targetMonth, 구분: "매장", Target: storeAmount },
    { 기준월: targetMonth, 구분: "비매장", Target: nonStoreAmount },
    {
      기준월: targetMonth,
      구분: "총 Target",
      Target: storeAmount + nonStoreAmount,
    },
  ];

  return (
    <div className="rounded-xl border border-gray-300/70 bg-white/80 p-4 shadow-sm backdrop-blur">
      <div className="mb-3 flex flex-col gap-2 lg:flex-row lg:items-center lg:justify-between">
        <div>
          <h2 className="text-lg font-bold">Target 관리</h2>
          <p className="mt-1 text-sm text-slate-500">
            Target은 거래처별이 아니라 매장 / 비매장 기준으로만 관리합니다.
          </p>
        </div>
        <div className="flex flex-wrap gap-2">
          <button
            onClick={() => exportExcel(excelRows, `Target관리_${targetMonth}`)}
            className="rounded-xl bg-emerald-600 px-4 py-2 text-sm font-semibold text-white shadow-sm hover:bg-emerald-700"
          >
            엑셀 다운로드
          </button>
          <label className="cursor-pointer rounded-lg bg-green-600 px-3 py-1.5 text-xs font-semibold text-white">
            엑셀 업로드
            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              className="hidden"
              onChange={(e) => upload(e.target.files?.[0] || null)}
            />
          </label>
        </div>
      </div>

      <div className="mb-4 flex flex-wrap gap-2">
        <input
          type="month"
          value={targetMonth}
          onChange={(e) => setTargetMonth(e.target.value)}
          className="rounded-xl border px-3 py-2"
        />
      </div>

      <div className="grid grid-cols-1 gap-4 lg:grid-cols-4">
        <div className="rounded-2xl border border-blue-100 bg-blue-50/70 p-4">
          <p className="text-sm font-semibold text-slate-600">매장 Target</p>
          <input
            value={storeAmount ? won(storeAmount) : ""}
            onChange={(e) => updateAmount("매장", num(e.target.value))}
            className="mt-3 w-full rounded-xl border border-slate-300 bg-white px-3 py-3 text-right text-xl font-bold text-slate-900 outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
            placeholder="0"
          />
        </div>
        <div className="rounded-2xl border border-amber-100 bg-amber-50/70 p-4">
          <p className="text-sm font-semibold text-slate-600">비매장 Target</p>
          <input
            value={nonStoreAmount ? won(nonStoreAmount) : ""}
            onChange={(e) => updateAmount("비매장", num(e.target.value))}
            className="mt-3 w-full rounded-xl border border-slate-300 bg-white px-3 py-3 text-right text-xl font-bold text-slate-900 outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
            placeholder="0"
          />
        </div>
        <div className="rounded-2xl border border-gray-300 bg-slate-50 p-4">
          <p className="text-sm font-semibold text-slate-600">총 Target</p>
          <p className="mt-3 break-all text-right text-2xl font-bold text-slate-900">
            {won(storeAmount + nonStoreAmount)}
          </p>
        </div>
      </div>
    </div>
  );
}

function TargetOrEstPage({
  title,
  records,
  setRecords,
  stores,
  month,
}: {
  title: string;
  records: TargetRecord[] | EstRecord[];
  setRecords: (v: any[]) => void;
  stores: Store[];
  month: string;
}) {
  const [targetMonth, setTargetMonth] = useState(month);
  const [search, setSearch] = useState("");
  const [searchDraft, setSearchDraft] = useState("");
  const isEstPage = title.includes("EST");
  const stMap = storeMap(stores);

  const rows = stores
    .filter((s) => s.status === "거래중")
    .filter((s) =>
      `${s.code} ${s.name} ${s.manager} ${s.channel}`
        .toLowerCase()
        .includes(search.toLowerCase()),
    )
    .map((s) => {
      const existing = records.find(
        (r) => r.storeCode === s.code && r.month === targetMonth,
      );
      return { store: s, amount: existing?.amount || 0 };
    });

  const estSummary = records
    .filter((r) => r.month === targetMonth)
    .reduce(
      (acc, r) => {
        const type =
          stMap.get(r.storeCode || "")?.storeType === "매장"
            ? "매장"
            : "비매장";
        if (type === "매장") acc.store += r.amount || 0;
        else acc.nonStore += r.amount || 0;
        return acc;
      },
      { store: 0, nonStore: 0 },
    );

  function resetEst() {
    if (!confirm(`${targetMonth} EST 데이터를 초기화할까요?`)) return;
    setRecords(records.filter((r) => r.month !== targetMonth));
    alert(`${targetMonth} EST 데이터가 초기화되었습니다.`);
  }

  function updateAmount(store: Store, amount: number) {
    const next = {
      storeCode: store.code,
      storeName: store.name,
      month: targetMonth,
      amount,
    };
    setRecords([
      ...records.filter(
        (r) => !(r.storeCode === store.code && r.month === targetMonth),
      ),
      next,
    ]);
  }

  async function upload(file: File | null) {
    if (!file) return;
    const fileRows = await readFileRows(file);
    const parsed = fileRows
      .map((r) => ({
        storeCode: norm(r["거래처코드"] ?? r["매장코드"]),
        storeName: norm(r["거래처명"] ?? r["매장명"]),
        month: monthText(r["기준월"] ?? r["월"] ?? r["년월"] ?? targetMonth),
        amount: num(
          r["당월 EST"] ??
            r["EST"] ??
            r["금액"] ??
            r["Target"] ??
            r["TARGET"] ??
            r["목표"] ??
            r["목표매출"],
        ),
      }))
      .filter((r) => r.storeCode && r.month && r.amount > 0);

    setRecords([
      ...records.filter(
        (r) =>
          !parsed.some(
            (p) => p.storeCode === r.storeCode && p.month === r.month,
          ),
      ),
      ...parsed,
    ]);
    alert(`${title} ${parsed.length}건을 반영했습니다.`);
  }

  return (
    <div className="rounded-xl border border-gray-300/70 bg-white/80 p-4 shadow-sm backdrop-blur">
      <div className="mb-3 flex flex-col gap-2 lg:flex-row lg:items-center lg:justify-between">
        <div>
          <h2 className="text-lg font-bold">{title}</h2>
          <p className="mt-1 text-sm text-slate-500">
            거래처별 {title.replace(" 관리", "")} 금액을 관리합니다.
          </p>
        </div>
        <div className="flex flex-wrap gap-2">
          <label className="cursor-pointer rounded-lg bg-green-600 px-3 py-1.5 text-xs font-semibold text-white">
            엑셀 업로드
            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              className="hidden"
              onChange={(e) => upload(e.target.files?.[0] || null)}
            />
          </label>
          {isEstPage && (
            <button
              type="button"
              onClick={resetEst}
              className="rounded-lg border border-red-200 bg-red-50 px-3 py-1.5 text-xs font-semibold text-red-600 hover:bg-red-100"
            >
              EST 리셋
            </button>
          )}
        </div>
      </div>

      {isEstPage && (
        <div className="mb-4 grid grid-cols-1 gap-3 lg:grid-cols-3">
          <div className="rounded-2xl border border-orange-200 bg-orange-50 p-4 shadow-sm">
            <p className="text-sm font-semibold text-slate-600">매장 EST</p>
            <p className="mt-2 break-all text-right text-2xl font-bold text-slate-900">
              {won(estSummary.store)}
            </p>
          </div>
          <div className="rounded-2xl border border-orange-200 bg-orange-50 p-4 shadow-sm">
            <p className="text-sm font-semibold text-slate-600">비매장 EST</p>
            <p className="mt-2 break-all text-right text-2xl font-bold text-slate-900">
              {won(estSummary.nonStore)}
            </p>
          </div>
          <div className="rounded-2xl border border-orange-300 bg-orange-100 p-4 shadow-sm">
            <p className="text-sm font-semibold text-slate-600">총 EST</p>
            <p className="mt-2 break-all text-right text-2xl font-bold text-slate-900">
              {won(estSummary.store + estSummary.nonStore)}
            </p>
          </div>
        </div>
      )}

      <div className="mb-4 flex flex-wrap gap-2">
        <input
          type="month"
          value={targetMonth}
          onChange={(e) => setTargetMonth(e.target.value)}
          className="rounded-xl border px-3 py-2"
        />
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

function buildAutoClosedStoresFromPrevYear(
  parsed: SalesRecord[],
  stores: Store[],
) {
  const existingByCode = new Map(stores.map((s) => [norm(s.code), s]));
  const currentActiveByName = new Set(
    stores
      .filter((s) => s.status === "거래중")
      .map((s) => normalizeStoreNameKey(s.name))
      .filter(Boolean),
  );
  const map = new Map<string, Store>();

  parsed.forEach((r) => {
    const codeKey = norm(r.storeCode);
    const nameKey = normalizeStoreNameKey(r.storeName);
    if (
      !codeKey ||
      existingByCode.has(codeKey) ||
      (nameKey && currentActiveByName.has(nameKey))
    )
      return;

    map.set(codeKey, {
      code: r.storeCode,
      name: r.storeName || r.storeCode,
      channel: r.channel || "미지정",
      manager: (r.manager || "") as Manager,
      storeType: r.storeType || "비매장",
      brand: displayBrand(r.brand),
      status: "거래종료",
    });
  });

  return Array.from(map.values());
}

function UploadPage({
  stores,
  setStores,
  sales,
  setSales,
  salesActions,
  month,
  date,
  timeConfigs,
  setTimeConfigs,
  itemMasters,
  setItemMasters,
}: {
  stores: Store[];
  setStores: (v: Store[]) => void;
  sales: SalesRecord[];
  setSales: React.Dispatch<React.SetStateAction<SalesRecord[]>>;
  salesActions: SalesStorageActions;
  month: string;
  date: string;
  timeConfigs: TimeConfig[];
  setTimeConfigs: (v: TimeConfig[]) => void;
  itemMasters: ItemMasterRecord[];
  setItemMasters: (v: ItemMasterRecord[]) => void;
}) {
  const [holidayText, setHolidayText] = useState("");
  const [deleteDate, setDeleteDate] = useState(today());

  async function uploadSales(file: File | null, period: PeriodType) {
    if (!file) return;
    const rows = await readFileRows(file);
    const parsed = rows
      .map((r, index) => {
        const saleDate = dateText(
          r["일자 No."] ??
            r["일자"] ??
            r["매출일"] ??
            r["판매일"] ??
            r["기준일"] ??
            deleteDate,
        );
        const storeCode = norm(
          r["거래처 코드"] ?? r["거래처코드"] ?? r["매장코드"],
        );
        const storeName = norm(r["거래처명"] ?? r["매장명"]);
        const itemCode = norm(r["품목 코드"] ?? r["품목코드"] ?? r["상품코드"]);
        const itemName = norm(r["품목명[규격]"] ?? r["품목명"] ?? r["상품명"]);
        const quantity = num(r["판매 수량"] ?? r["수량"]);
        const rawSalesAmount =
          r["판매 금액"] ??
          r["판매금액"] ??
          r["매출금액"] ??
          r["당월 매출"] ??
          r["매출"];
        const rawCostAmount = r["원가 금액"] ?? r["원가금액"] ?? r["원가"];
        const rawProfitAmount =
          r["이익 금액"] ??
          r["이익금액"] ??
          r["매출 이익"] ??
          r["매출이익"] ??
          r["매출총이익"] ??
          r["마진액"];
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
            r["마진율"],
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
          uploadedProfitRate,
        );
      })
      .filter(
        (r) =>
          r.saleDate &&
          r.storeCode &&
          (r.salesAmount !== 0 ||
            r.costAmount !== 0 ||
            r.profitAmount !== 0 ||
            r.quantity !== 0),
      );

    const missingStores =
      period === "current"
        ? parsed
            .filter((r) => !storeMap(stores).has(r.storeCode))
            .map((r) => ({
              code: r.storeCode,
              name: r.storeName || r.storeCode,
              channel: "매장" as Channel,
              manager: "" as Manager,
              storeType: "매장" as StoreType,
              brand: displayBrand(r.brand),
              status: "거래중" as const,
            }))
        : [];

    if (missingStores.length) {
      const map = new Map(stores.map((s) => [s.code, s]));
      missingStores.forEach((s) => map.set(s.code, s));
      setStores(Array.from(map.values()));
    }

    if (parsed.length === 0) {
      alert(
        "업로드 파일에서 반영할 매출/원가/이익 행을 찾지 못했습니다. 기존 데이터는 삭제하지 않았습니다.",
      );
      return;
    }

    const uploadedDates = Array.from(
      new Set(parsed.map((r) => r.saleDate).filter(Boolean)),
    );

    try {
      const result = await salesActions.replaceUpload({
        period,
        refMonth: month,
        fileName: file.name,
        uploadedDates,
        rows: parsed,
      });
      if (result.mode === "legacy") {
        console.warn("Sales V3 SQL이 아직 적용되지 않아 기존 저장 방식을 사용했습니다.");
      }
    } catch (error) {
      console.error("매출 업로드 저장 실패", error);
      alert("매출 저장에 실패했습니다. 기존 데이터는 그대로 유지됩니다. Cloudflare D1 연결 상태를 확인해 주세요.");
      return;
    }

    const closedMessage = "";
    alert(
      `${period === "current" ? "당월" : period === "prevMonth" ? "전월" : "전년동월"} 매출 ${parsed.length}건을 반영했습니다.\n반영 날짜: ${uploadedDates.join(", ")}${closedMessage}`,
    );
  }

  async function uploadItems(file: File | null) {
    if (!file) return;
    const rows = await readFileRows(file);
    const parsed = rows
      .map((row) => ({
        itemCode: norm(
          row["품목코드"] ?? row["품목 코드"] ?? row["상품코드"],
        ),
        itemName: norm(
          row["품목명"] ?? row["품목명[규격]"] ?? row["상품명"],
        ),
        category: norm(
          row["카테고리"] ?? row["품목그룹"] ?? row["분류"],
        ),
        active:
          !["중단", "미사용", "N", "FALSE", "0"].includes(
            norm(row["사용여부"] ?? row["사용 여부"]).toUpperCase(),
          ),
        memo: norm(row["메모"] ?? row["비고"]),
      }))
      .filter((row) => row.itemCode && row.category);

    if (!parsed.length) {
      alert(
        "품목코드와 카테고리가 있는 행을 찾지 못했습니다. 파일 헤더를 확인해 주세요.",
      );
      return;
    }

    const map = new Map(itemMasters.map((item) => [item.itemCode, item]));
    parsed.forEach((item) => {
      const previous = map.get(item.itemCode);
      map.set(item.itemCode, {
        ...previous,
        ...item,
        itemName: item.itemName || previous?.itemName || item.itemCode,
      });
    });
    setItemMasters(Array.from(map.values()));
    alert(`품목 카테고리 ${parsed.length}건을 반영했습니다.`);
  }

  function saveHolidays() {
    const holidays: string[] = Array.from(
      new Set<string>(
        holidayText
          .split(/\n|,|\s+/)
          .map(dateText)
          .filter((d): d is string => Boolean(d && d.startsWith(month))),
      ),
    ).sort();
    setTimeConfigs([
      ...timeConfigs.filter((c) => c.month !== month),
      { month, holidays },
    ]);
    alert(`${month} TIME GONE 공휴일 ${holidays.length}건을 저장했습니다.`);
  }

  async function deleteCurrentDate() {
    if (!confirm(`${deleteDate} 당월 매출 데이터를 삭제할까요?`)) return;
    try {
      await salesActions.deleteCurrentDate(month, deleteDate);
      alert(`${deleteDate} 당월 매출 데이터를 삭제 처리했습니다.`);
    } catch (error) {
      console.error("매출 삭제 실패", error);
      alert("삭제에 실패했습니다. 기존 데이터는 유지됩니다.");
    }
  }

  return (
    <div className="space-y-4">
      <div className="rounded-2xl border border-gray-300/70 bg-white/80 p-5 shadow-sm backdrop-blur">
        <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
          <h2 className="text-lg font-bold">매출 업로드</h2>
          <span className={`rounded-full px-3 py-1 text-xs font-bold ${salesActions.storageMode === "v3" ? "bg-emerald-100 text-emerald-800" : salesActions.storageMode === "checking" ? "bg-slate-100 text-slate-600" : "bg-amber-100 text-amber-800"}`}>
            {salesActions.storageMode === "v3" ? "안전 저장 V3" : salesActions.storageMode === "checking" ? "저장소 확인 중" : "기존 저장 방식"}
          </span>
        </div>
        <div className="grid grid-cols-1 gap-2 sm:grid-cols-2 xl:grid-cols-4">
          <UploadBox
            title="당월 매출 업로드"
            description="같은 날짜 파일을 다시 올리면 해당 날짜 기존 당월 매출을 삭제하고 새 파일로 업데이트합니다."
            onUpload={(file) => uploadSales(file, "current")}
          />
          <UploadBox
            title="전월 매출 업로드"
            description="매출비교와 매출현황의 전월 매출 기준으로 사용합니다. 같은 기준월 자료는 새 파일로 교체됩니다."
            onUpload={(file) => uploadSales(file, "prevMonth")}
          />
          <UploadBox
            title="전년동월 매출 업로드"
            description="매출비교와 매출현황의 전년동월 매출 기준으로 사용합니다. 같은 기준월 자료는 새 파일로 교체됩니다."
            onUpload={(file) => uploadSales(file, "prevYear")}
          />
          <UploadBox
            title="품목 정보 업로드"
            description="품목코드와 카테고리를 업로드하면 품목분석의 카테고리에 반영됩니다. 기존 품목은 갱신하고 파일에 없는 품목은 유지합니다."
            onUpload={uploadItems}
          />
        </div>
      </div>


      <div className="rounded-2xl border border-gray-300/70 bg-white/80 p-5 shadow-sm backdrop-blur">
        <h2 className="mb-3 text-lg font-bold">당월 특정 날짜 삭제</h2>
        <input
          type="date"
          value={deleteDate}
          onChange={(e) => setDeleteDate(e.target.value)}
          className="mr-2 rounded-xl border px-3 py-2"
        />
        <button
          onClick={deleteCurrentDate}
          className="rounded-xl bg-red-600 px-4 py-2 text-white"
        >
          해당일 당월 매출 삭제
        </button>
      </div>

      <div className="rounded-2xl border border-gray-300/70 bg-white/80 p-5 shadow-sm backdrop-blur">
        <h2 className="mb-3 text-lg font-bold">TIME GONE 공휴일 설정</h2>
        <p className="mb-3 text-sm text-slate-500">
          월~금 일반일 1일, 공휴일 0.5일, 토요일 0.5일, 일요일 0.5일
          기준입니다.
        </p>
        <textarea
          value={holidayText}
          onChange={(e) => setHolidayText(e.target.value)}
          placeholder={`${month}-06\n${month}-15`}
          className="h-28 w-full rounded-xl border px-3 py-2"
        />
        <button
          onClick={saveHolidays}
          className="mt-3 rounded-xl bg-blue-600 px-4 py-2 text-white"
        >
          공휴일 저장
        </button>
      </div>
    </div>
  );
}

function ProfitValidationPanel({
  sales,
  month,
  date,
}: {
  sales: SalesRecord[];
  month: string;
  date: string;
}) {
  const monthRows = sales.filter(
    (s) =>
      s.period === "current" &&
      inRange(s.saleDate, monthStart(month), monthEnd(month)),
  );
  const toDateRows = sales.filter(
    (s) =>
      s.period === "current" && inRange(s.saleDate, monthStart(month), date),
  );
  const monthProfit = sum(monthRows, "profitAmount");
  const toDateProfit = sum(toDateRows, "profitAmount");
  const monthSales = sum(monthRows, "salesAmount");
  const toDateSales = sum(toDateRows, "salesAmount");
  const monthRate = weightedProfitRate(monthRows);
  const toDateRate = weightedProfitRate(toDateRows);

  const dailyRows: {
    date: string;
    salesAmount: number;
    profitAmount: number;
    profitRate: number;
  }[] = [];
  for (
    let d = monthStart(month), guard = 0;
    d <= monthEnd(month) && guard < 40;
    d = addDays(d, 1), guard += 1
  ) {
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
    <div className="rounded-2xl border border-gray-300/70 bg-white/80 p-5 shadow-sm backdrop-blur">
      <div className="mb-3 flex flex-col gap-2 lg:flex-row lg:items-center lg:justify-between">
        <div>
          <h2 className="text-lg font-bold">이익금액 검증</h2>
          <p className="mt-1 text-sm text-slate-500">
            당월 전체와 기준일({date})까지의 이익금액을 나누어 확인합니다.
          </p>
        </div>
        <button
          type="button"
          onClick={() =>
            exportExcel(excelRows, `이익금액_날짜별_검증_${month}`)
          }
          className="rounded-xl bg-emerald-600 px-4 py-2 text-sm font-semibold text-white shadow-sm hover:bg-emerald-700"
        >
          날짜별 검증표 다운로드
        </button>
      </div>

      <div className="mb-4 grid grid-cols-1 gap-3 md:grid-cols-2">
        <div className="rounded-xl border border-green-200 bg-green-50 p-4">
          <p className="text-sm font-semibold text-slate-600">당월 이익금액</p>
          <p className="mt-2 text-2xl font-bold text-slate-900">
            {won(monthProfit)}원
          </p>
          <p className="mt-1 text-xs text-slate-500">
            당월 매출 {won(monthSales)}원 · 이익률 {pct(monthRate)}
          </p>
        </div>
        <div className="rounded-xl border border-amber-200 bg-amber-50 p-4">
          <p className="text-sm font-semibold text-slate-600">
            당일까지의 이익
          </p>
          <p className="mt-2 text-2xl font-bold text-slate-900">
            {won(toDateProfit)}원
          </p>
          <p className="mt-1 text-xs text-slate-500">
            당일까지 매출 {won(toDateSales)}원 · 이익률 {pct(toDateRate)}
          </p>
        </div>
      </div>

      <div className="max-h-[360px] overflow-auto">
        <table className="w-full min-w-[640px] border border-gray-300 text-xs">
          <thead>
            <tr className="bg-slate-100">
              <th className="sticky top-0 border bg-slate-100 px-3 py-2 text-left font-bold">
                날짜
              </th>
              <th className="sticky top-0 border bg-slate-100 px-3 py-2 text-right font-bold">
                매출금액
              </th>
              <th className="sticky top-0 border bg-slate-100 px-3 py-2 text-right font-bold">
                이익금액
              </th>
              <th className="sticky top-0 border bg-slate-100 px-3 py-2 text-right font-bold">
                이익률
              </th>
            </tr>
          </thead>
          <tbody>
            {dailyRows.map((r) => (
              <tr
                key={r.date}
                className={
                  r.date <= date ? "bg-white" : "bg-slate-50 text-slate-500"
                }
              >
                <td className="border px-3 py-2 font-semibold">{r.date}</td>
                <td className="border px-3 py-2 text-right font-semibold text-slate-900">
                  {won(r.salesAmount)}
                </td>
                <td className="border px-3 py-2 text-right font-semibold text-slate-900">
                  {won(r.profitAmount)}
                </td>
                <td className="border px-3 py-2 text-right">
                  {pct(r.profitRate)}
                </td>
              </tr>
            ))}
          </tbody>
          <tfoot>
            <tr className="bg-slate-100 font-bold">
              <td className="border px-3 py-2">합계</td>
              <td className="border px-3 py-2 text-right">{won(monthSales)}</td>
              <td className="border px-3 py-2 text-right">
                {won(monthProfit)}
              </td>
              <td className="border px-3 py-2 text-right">{pct(monthRate)}</td>
            </tr>
          </tfoot>
        </table>
      </div>
    </div>
  );
}

function UploadBox({
  title,
  description,
  onUpload,
}: {
  title: string;
  description: string;
  onUpload: (file: File | null) => void;
}) {
  return (
    <div className="flex min-h-[116px] flex-col rounded-xl border border-gray-300 bg-slate-50/70 p-3">
      <h3 className="text-sm font-bold text-slate-900">{title}</h3>
      <p className="mt-1 line-clamp-2 flex-1 text-xs leading-5 text-slate-500">{description}</p>
      <label className="mt-2 inline-flex w-full cursor-pointer items-center justify-center rounded-lg bg-blue-600 px-3 py-1.5 text-xs font-semibold text-white hover:bg-blue-700">
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

function Th({
  children,
  right = false,
  tone = "default",
}: {
  children: React.ReactNode;
  right?: boolean;
  tone?: "default" | "mint" | "blue" | "pink" | "yellow" | "gray" | "purple" | "green" | "orange";
}) {
  const toneClass =
    tone === "mint"
      ? "border-[#E5E7EB] bg-[#F7FCEB] text-black"
      : tone === "blue"
        ? "border-[#E5E7EB] bg-[#F3FAFD] text-black"
        : tone === "pink"
          ? "border-[#E5E7EB] bg-[#FFF7FA] text-black"
          : tone === "yellow"
            ? "border-[#E5E7EB] bg-[#FFFDF2] text-black"
            : tone === "orange"
              ? "border-[#E5E7EB] bg-[#FFF9F3] text-black"
              : "border-slate-300 bg-white text-black";

  return (
    <th
      className={`sticky top-0 z-20 border p-2 text-xs font-bold leading-tight whitespace-normal break-keep ${toneClass} text-center`}
    >
      {children}
    </th>
  );
}

function Td({
  children,
  right = false,
  bold = false,
  color = "",
}: {
  children: React.ReactNode;
  right?: boolean;
  bold?: boolean;
  color?: string;
}) {
  return (
    <td
      className={`border border-gray-300 bg-white p-2 text-center ${bold ? "font-semibold" : ""} ${color}`}
    >
      {children}
    </td>
  );
}
