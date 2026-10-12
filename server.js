import express from 'express';
import cors from 'cors';
import multer from 'multer';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';
import XLSX from 'xlsx';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3000;

// Directories
const DATA_DIR = path.join(__dirname, 'data');
const UPLOADS_DIR = path.join(__dirname, 'uploads');
const TEMPLATES_DIR = path.join(UPLOADS_DIR, 'templates');
const PUBLIC_DIR = path.join(__dirname, 'public');

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
if (!fs.existsSync(UPLOADS_DIR)) fs.mkdirSync(UPLOADS_DIR, { recursive: true });
if (!fs.existsSync(TEMPLATES_DIR)) fs.mkdirSync(TEMPLATES_DIR, { recursive: true });
if (!fs.existsSync(PUBLIC_DIR)) fs.mkdirSync(PUBLIC_DIR, { recursive: true });

app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(PUBLIC_DIR));
app.use('/uploads', express.static(UPLOADS_DIR));

// Helper: JSON File Store
const getFilePath = (fileName) => path.join(DATA_DIR, fileName);

const readJSON = (fileName, defaultVal = []) => {
  const file = getFilePath(fileName);
  if (!fs.existsSync(file)) {
    fs.writeFileSync(file, JSON.stringify(defaultVal, null, 2), 'utf-8');
    return defaultVal;
  }
  try {
    const data = fs.readFileSync(file, 'utf-8');
    return JSON.parse(data);
  } catch (err) {
    console.error(`Error reading ${fileName}:`, err);
    return defaultVal;
  }
};

const writeJSON = (fileName, data) => {
  const file = getFilePath(fileName);
  fs.writeFileSync(file, JSON.stringify(data, null, 2), 'utf-8');
};

// Multer storage for contracts
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, UPLOADS_DIR);
  },
  filename: (req, file, cb) => {
    const ext = path.extname(file.originalname);
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1e9);
    cb(null, `${uniqueSuffix}_${file.originalname}`);
  }
});
const upload = multer({ storage });

// Multer storage for Process Guide Templates
const templateStorage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, TEMPLATES_DIR);
  },
  filename: (req, file, cb) => {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1e9);
    cb(null, `template_${uniqueSuffix}_${file.originalname}`);
  }
});
const uploadTemplate = multer({ storage: templateStorage });

// Initial Seed Data
const initData = () => {
  // 1. Projects
  const projects = readJSON('projects.json', null);
  if (!projects || projects.length === 0) {
    const seedProjects = [
      {
        id: 'P-2026-001',
        year: 2026,
        code: 'PRJ-2026-01',
        name: 'AI 기반 공공정책 데이터 분석 고도화 연구',
        department: '디지털혁신연구실',
        manager: '김철수 책임연구원',
        email: 'chulsoo.kim@re.kr',
        budget: 150000000,
        startDate: '2026-01-01',
        endDate: '2026-12-31',
        description: '공공정책 수립을 위한 초거대 AI 기반 분석 모델 개발'
      },
      {
        id: 'P-2026-002',
        year: 2026,
        code: 'PRJ-2026-02',
        name: '국가 균형발전 지표 조사 및 영향도 분석',
        department: '지역발전연구본부',
        manager: '이영희 선임연구원',
        email: 'younghee.lee@re.kr',
        budget: 80000000,
        startDate: '2026-02-01',
        endDate: '2026-10-31',
        description: '전국 17개 시도 균형발전 지표 대국민 설문조사 및 정밀 실태조사'
      },
      {
        id: 'P-2026-003',
        year: 2026,
        code: 'PRJ-2026-03',
        name: '2026 경제인문사회연구회 협동연구 : 미래 인구구조 대응 전략',
        department: '사회정책연구실',
        manager: '박민수 연구위원',
        email: 'minsu.park@re.kr',
        budget: 250000000,
        startDate: '2026-01-15',
        endDate: '2026-11-30',
        description: '경소연 협동연구과제 수행을 위한 타 연구기관 및 학계 공동협력 연구'
      },
      {
        id: 'P-2025-001',
        year: 2025,
        code: 'PRJ-2025-05',
        name: '탄소중립 이행을 위한 에너지 전환 정책 연구',
        department: '환경경제연구실',
        manager: '정다은 책임연구원',
        email: 'daeun.jung@re.kr',
        budget: 120000000,
        startDate: '2025-01-01',
        endDate: '2025-12-31',
        description: '에너지 전환 산업계 실태조사'
      }
    ];
    writeJSON('projects.json', seedProjects);
  }

  // 2. Contracts
  const contracts = readJSON('contracts.json', null);
  if (!contracts || contracts.length === 0) {
    const seedContracts = [
      {
        id: 'C-2026-001',
        year: 2026,
        projectId: 'P-2026-001',
        projectName: 'AI 기반 공공정책 데이터 분석 고도화 연구',
        contractName: 'AI 분석 알고리즘 고도화 및 데이터 정제 용역',
        contractType: '공개경쟁입찰',
        contractNature: '조사용역계약',
        contractAmount: 120000000,
        contractorName: '(주)한국인공지능시스템즈',
        businessNumber: '123-81-99887',
        representative: '홍길동',
        contractorContact: '02-555-0199 / contact@k-ai.co.kr',
        department: '디지털혁신연구실',
        manager: '김철수 책임연구원',
        managerEmail: 'chulsoo.kim@re.kr',
        startDate: '2026-02-01',
        endDate: '2026-08-31',
        advancePayment: {
          hasAdvance: true,
          amount: 40000000,
          date: '2026-02-15',
          insuranceEndDate: '2026-10-30',
          planSubmitted: true
        },
        progressStep: 5, // 1~6단계
        status: '준공검수', // 계획기안, 협조문작성, 계약체결, 착수진행, 준공검수, 대금기안, 지출완료
        completionDate: '2026-08-31',
        inspectionDate: '',
        inspectionDueDate: '2026-09-14',
        inspectionSubmitted: false,
        paymentDraftDate: '',
        paymentDueDate: '',
        paymentPaidDate: '',
        isPaid: false,
        documents: [
          { id: 'doc-1', category: '계약체결', title: '계약서 원본', fileName: '계약서_AI고도화.pdf', uploadedAt: '2026-02-01', isDefault: true, required: true },
          { id: 'doc-2', category: '계약체결', title: '보안서약서', fileName: '보안서약서_서명본.pdf', uploadedAt: '2026-02-01', isDefault: true, required: true },
          { id: 'doc-3', category: '계약체결', title: '계약보증보험증권', fileName: '계약이행보증증권.pdf', uploadedAt: '2026-02-01', isDefault: true, required: true },
          { id: 'doc-4', category: '선급금', title: '선금보증보험증권 (종료일+60일)', fileName: '선금보증보험증권.pdf', uploadedAt: '2026-02-15', isDefault: true, required: true },
          { id: 'doc-5', category: '선급금', title: '선금사용계획서', fileName: '선금사용계획서_승인본.pdf', uploadedAt: '2026-02-15', isDefault: true, required: true }
        ],
        notes: '공개경쟁입찰 15일 공고 완료 및 기술평가위원회 진행 완료 건'
      },
      {
        id: 'C-2026-002',
        year: 2026,
        projectId: 'P-2026-002',
        projectName: '국가 균형발전 지표 조사 및 영향도 분석',
        contractName: '전국 17개 시도 균형발전 인식 실태 설문조사 용역',
        contractType: '수의계약',
        contractNature: '조사용역계약',
        contractAmount: 35000000,
        contractorName: '(주)글로벌리서치센터',
        businessNumber: '214-86-12345',
        representative: '강대성',
        contractorContact: '02-789-0123 / survey@globalresearch.com',
        department: '지역발전연구본부',
        manager: '이영희 선임연구원',
        managerEmail: 'younghee.lee@re.kr',
        startDate: '2026-02-15',
        endDate: '2026-04-15',
        advancePayment: { hasAdvance: false },
        progressStep: 5,
        status: '준공검수',
        completionDate: '2026-04-15',
        inspectionDueDate: '2026-04-29',
        inspectionDate: '',
        inspectionSubmitted: false,
        paymentDraftDate: '',
        paymentDueDate: '',
        paymentPaidDate: '',
        isPaid: false,
        documents: [
          { id: 'doc-201', category: '계약체결', title: '계약서 원본', fileName: '용역계약서_설문조사.pdf', uploadedAt: '2026-02-15', isDefault: true, required: true },
          { id: 'doc-202', category: '계약체결', title: '견적서 (주업체)', fileName: '글로벌리서치_견적서.pdf', uploadedAt: '2026-02-14', isDefault: true, required: true },
          { id: 'doc-203', category: '계약체결', title: '비교견적서 (타업체 1종 이상)', fileName: '한국통계정보_비교견적서.pdf', uploadedAt: '2026-02-14', isDefault: true, required: true },
          { id: 'doc-204', category: '준공시점', title: '준공계', fileName: '준공계_글로벌리서치.pdf', uploadedAt: '2026-04-15', isDefault: true, required: true },
          { id: 'doc-205', category: '준공시점', title: '대금청구서', fileName: '세금계산서_및_대금청구서.pdf', uploadedAt: '2026-04-15', isDefault: true, required: true }
        ],
        notes: '비교견적 2종 첨부 수의계약. 현재 14일 이내 검수보고서 수령 대기 중 (알림 대상)'
      },
      {
        id: 'C-2026-003',
        year: 2026,
        projectId: 'P-2026-003',
        projectName: '2026 경제인문사회연구회 협동연구 : 미래 인구구조 대응 전략',
        contractName: '인구동태 예측 모델링 위탁연구 용역',
        contractType: '수의계약',
        contractNature: '협동연구계약',
        contractAmount: 48000000,
        contractorName: '서울대학교 산학협력단',
        businessNumber: '119-82-00312',
        representative: '유홍림',
        contractorContact: '02-880-1114 / snu_rnd@snu.ac.kr',
        department: '사회정책연구실',
        manager: '박민수 연구위원',
        managerEmail: 'minsu.park@re.kr',
        startDate: '2026-03-01',
        endDate: '2026-09-30',
        advancePayment: { hasAdvance: false },
        progressStep: 3,
        status: '계약체결',
        completionDate: '2026-09-30',
        inspectionDueDate: '2026-10-14',
        inspectionSubmitted: false,
        paymentDraftDate: '',
        paymentDueDate: '',
        paymentPaidDate: '',
        isPaid: false,
        documents: [
          { id: 'doc-301', category: '계약체결', title: '협동연구계약서', fileName: '협동연구계약서_서울대.pdf', uploadedAt: '2026-03-01', isDefault: true, required: true },
          { id: 'doc-302', category: '계약체결', title: '연구계획서 및 견적서', fileName: '위탁연구계획서_견적서.pdf', uploadedAt: '2026-03-01', isDefault: true, required: true }
        ],
        notes: '경소연 협동연구과제 수행을 위한 대학 협동연구 계약'
      },
      {
        id: 'C-2026-004',
        year: 2026,
        projectId: 'P-2026-001',
        projectName: 'AI 기반 공공정책 데이터 분석 고도화 연구',
        contractName: '정책자문 및 공동연구 자문 (최외부 박사)',
        contractType: '수의계약',
        contractNature: '공동연구진계약',
        contractAmount: 4500000,
        contractorName: '최외부 박사 (한국데이터학회)',
        businessNumber: '750312-1******',
        representative: '최외부',
        contractorContact: '010-9988-7766 / dr.choi@data.org',
        department: '디지털혁신연구실',
        manager: '김철수 책임연구원',
        managerEmail: 'chulsoo.kim@re.kr',
        startDate: '2026-03-10',
        endDate: '2026-05-31',
        advancePayment: { hasAdvance: false },
        progressStep: 6,
        status: '대금기안',
        completionDate: '2026-05-31',
        inspectionDate: '2026-06-05',
        inspectionDueDate: '2026-06-14',
        inspectionSubmitted: true,
        paymentDraftDate: '2026-06-07',
        paymentDueDate: '2026-06-12',
        paymentPaidDate: '',
        isPaid: false,
        documents: [
          { id: 'doc-401', category: '계약체결', title: '공동연구진 계약서', fileName: '공동연구진_계약서.pdf', uploadedAt: '2026-03-10', isDefault: true, required: true },
          { id: 'doc-402', category: '계약체결', title: '연구심의(연심) 결과서', fileName: '연심결과_최외부박사.pdf', uploadedAt: '2026-03-10', isDefault: true, required: true },
          { id: 'doc-403', category: '계약체결', title: '단일견적서 (500만원 이하)', fileName: '전문가_단일견적서.pdf', uploadedAt: '2026-03-10', isDefault: true, required: true },
          { id: 'doc-404', category: '준공시점', title: '자문보고서 및 청구서', fileName: '최종자문결과보고서.pdf', uploadedAt: '2026-05-31', isDefault: true, required: true },
          { id: 'doc-405', category: '검수시점', title: '검수보고서', fileName: '검수보고서_승인.pdf', uploadedAt: '2026-06-05', isDefault: true, required: true }
        ],
        notes: '5백만원 이하 단일견적 수의계약. 검수보고서 접수 후 대금기안 완료 -> 5일 이내 지출 대기 중'
      }
    ];
    writeJSON('contracts.json', seedContracts);
  }

  // 3. Process Guide Steps & Default Templates
  const processGuide = readJSON('process_guide.json', null);
  if (!processGuide || processGuide.length === 0) {
    const seedGuide = [
      {
        id: 'step-1',
        stepNum: '1',
        title: '계약에 대한 최초 계획기안',
        department: '해당 연구부서',
        description: '- 각종 조사용역, 공동연구진, 기타 등의 계약을 통해 진행해야하는 업무의 최초 계획 기안',
        notes: '- 주의사항 : 업체견적서 등이 포함되면 안 됨\n(계약전 사전발주로 감사지적 사항임-사업부서에서의 발주는 금지사항이며 계약의 발주는 계약부서만 할 수 있음)',
        isDangerNote: true,
        templates: [
          { id: 't-1', name: '표준 계약 계획기안문 서식', fileName: '표준계획기안_서식.hwp', filePath: '/uploads/templates/표준계획기안_서식.hwp', uploadedAt: '2026-01-01' }
        ]
      },
      {
        id: 'step-2',
        stepNum: '2',
        title: '계약체결 협조문 작성',
        department: '해당 연구부서',
        description: '- 계약업무 진행을 계약부서로 협조요청',
        notes: '- 팀주무는 협조문 발송처리 확인필요 (발송지연으로 인한 계약체결지연 방지 목적)',
        isDangerNote: false,
        templates: [
          { id: 't-2', name: '계약체결 협조요청문(표준 양식)', fileName: '계약체결협조문_양식.hwp', filePath: '/uploads/templates/계약체결협조문_양식.hwp', uploadedAt: '2026-01-01' }
        ]
      },
      {
        id: 'step-3',
        stepNum: '3',
        title: '계약체결 기안 (수의계약일 경우)',
        department: '회계정보팀',
        description: '- 협조문을 근거로 하여 계약체결 기안 작성',
        notes: '- 계약서, 보안서약서 및 계약보증보험증권(계약보증지급각서) 등 기타 계약관련 서류 징구\n- 500만원 초과~1억원 미만 : 비교견적서 2종 이상 징구 필수\n- 500만원 이하 : 단일견적서로 수의계약 체결 가능',
        isDangerNote: false,
        templates: [
          { id: 't-3-1', name: '표준 용역계약서 서식', fileName: '표준용역계약서.docx', filePath: '/uploads/templates/표준용역계약서.docx', uploadedAt: '2026-01-01' },
          { id: 't-3-2', name: '보안서약서 표준 양식', fileName: '보안서약서_양식.hwp', filePath: '/uploads/templates/보안서약서_양식.hwp', uploadedAt: '2026-01-01' },
          { id: 't-3-3', name: '계약보증지급각서 서식', fileName: '계약보증지급각서.hwp', filePath: '/uploads/templates/계약보증지급각서.hwp', uploadedAt: '2026-01-01' }
        ]
      },
      {
        id: 'step-3-1',
        stepNum: '3-1',
        title: '입찰공고 기안 (공개경쟁입찰일 경우)',
        department: '회계정보팀',
        description: '- 입찰계약건(1억원 이상)일 경우 입찰공고 기안하여 홈페이지 및 알리오 등에 게시',
        notes: '- 제안요청서, 공고문 등 작성\n- 공고기간 15일',
        isDangerNote: false,
        templates: [
          { id: 't-3-1-1', name: '입찰공고문 표준 서식', fileName: '입찰공고문_표준서식.hwp', filePath: '/uploads/templates/입찰공고문_표준서식.hwp', uploadedAt: '2026-01-01' },
          { id: 't-3-1-2', name: '제안요청서(RFP) 표준 템플릿', fileName: '제안요청서_표준양식.docx', filePath: '/uploads/templates/제안요청서_표준양식.docx', uploadedAt: '2026-01-01' }
        ]
      },
      {
        id: 'step-3-1-1',
        stepNum: '3-1-1',
        title: '입찰재공고 기안',
        department: '회계정보팀',
        description: '- 1차 입찰건에 입찰참가 업체가 없거나 1개 업체일 경우 유찰되어 재공고함',
        notes: '- 공고기간 1주일\n- 재공고에도 유찰일 경우 해당 업체와 수의계약 체결\n- 한시적으로 유찰시 재공고 없이 수의계약 체결 가능(특례 규정 준용)',
        isDangerNote: false,
        templates: []
      },
      {
        id: 'step-3-2',
        stepNum: '3-2',
        title: '기술평가회 개최 & 결과보고 (3-2, 3-3)',
        department: '회계정보팀 / 해당사업부서',
        description: '- 입찰에 참가한 업체에 대한 기술평가회 개최 기안 및 평가결과 보고 후 선정업체와 계약체결 기안',
        notes: '- 각 업체에서 제출한 제안서와 평가표를 작성하여 기술평가회 실시\n- 계약체결시 징구서류는 3의 내용과 동일',
        isDangerNote: false,
        templates: [
          { id: 't-3-2-1', name: '기술평가표 양식', fileName: '기술제안서_평가표.xlsx', filePath: '/uploads/templates/기술제안서_평가표.xlsx', uploadedAt: '2026-01-01' }
        ]
      },
      {
        id: 'step-4',
        stepNum: '4',
        title: '선급금 지급 기안',
        department: '회계정보팀',
        description: '- 업체로부터 선급금 요청이 있을 경우 지급 기안 처리',
        notes: '- 선금보증보험증권, 선금사용계획서 등을 함께 수령해야 함\n- ⚠️ 주의: 선금보증보험의 보증기간은 반드시 [계약종료일 + 60일]로 해야 함',
        isDangerNote: true,
        templates: [
          { id: 't-4-1', name: '선금사용계획서 표준 서식', fileName: '선금사용계획서_양식.hwp', filePath: '/uploads/templates/선금사용계획서_양식.hwp', uploadedAt: '2026-01-01' },
          { id: 't-4-2', name: '선금지급신청서 양식', fileName: '선금지급신청서.hwp', filePath: '/uploads/templates/선금지급신청서.hwp', uploadedAt: '2026-01-01' }
        ]
      },
      {
        id: 'step-5',
        stepNum: '5',
        title: '용역대금 지급 & 검수',
        department: '해당사업부서(검수보고서 제출) / 총무정보팀',
        description: '- 계약 종료 시기에 업체로부터 준공계 및 대금청구서를 수령\n- 해당 연구부서로부터 "검수보고서"를 14일 이내 수령하여 이를 근거로 대금지급 기안',
        notes: '- 검수보고서는 제안요청서 및 계약사항을 확인하여 용역사업의 결과의 사실과 부합하게 작성\n- 🚨 기한 알림: 계약종료 후 14일 초과시 시스템 빨간색 강조 표시 및 담당자 이메일 발송',
        isDangerNote: true,
        templates: [
          { id: 't-5-1', name: '표준 검수보고서 양식 (14일 이내 필수)', fileName: '검수보고서_표준서식.hwp', filePath: '/uploads/templates/검수보고서_표준서식.hwp', uploadedAt: '2026-01-01' },
          { id: 't-5-2', name: '준공계(완료계) 표준 양식', fileName: '준공계_표준서식.hwp', filePath: '/uploads/templates/준공계_표준서식.hwp', uploadedAt: '2026-01-01' },
          { id: 't-5-3', name: '대금청구서 양식', fileName: '용역대금청구서.hwp', filePath: '/uploads/templates/용역대금청구서.hwp', uploadedAt: '2026-01-01' }
        ]
      },
      {
        id: 'step-6',
        stepNum: '6',
        title: '대금 지출 처리',
        department: '회계정보팀',
        description: '- 용역대금지급 기안 결재 후 청구일(기안일)로부터 5일 이내 지출 처리',
        notes: '- 🚨 기한 알림: 대금지급기안 후 5일 초과시 시스템 빨간색 강조 표시 및 회계팀 확인',
        isDangerNote: true,
        templates: [
          { id: 't-6-1', name: '대금지출결의서 서식', fileName: '대금지출결의서_양식.hwp', filePath: '/uploads/templates/대금지출결의서_양식.hwp', uploadedAt: '2026-01-01' }
        ]
      }
    ];
    writeJSON('process_guide.json', seedGuide);
  }

  // 4. Process Classification Rules (가, 나, 다)
  const processRules = readJSON('process_rules.json', null);
  if (!processRules) {
    const seedRules = {
      methodRules: {
        title: '가. 계약방법에 의한 분류',
        content: '1. 공개경쟁입찰계약\n- 우리 연구소 홈페이지 및 알리오, 기타 나라장터에 공개, 입찰로 진행하는 계약\n\n2. 수의계약\n- 국가를 당사자로 하는 계약에 관한 법률에 의거하여 계약을 체결하고자 하는 업체를 임의적으로 정하여 진행하는 계약의 형태'
      },
      amountRules: {
        title: '나. 계약금액별 체결방법',
        content: '• 계약금액 1억원 이상 : 공개경쟁입찰\n• 계약금액 1억원 미만 : 수의계약 체결 가능 (견적서 및 비교견적서 2종 이상 필요)\n※ 단, 5백만원 이하의 계약건은 단일견적으로 수의계약 가능'
      },
      natureRules: {
        title: '다. 계약성격에 의한 구분',
        content: '• 조사용역계약 : 과제를 수행하기 위한 각종 설문조사 용역계약\n• 공동연구진(연구협력진) 계약 : 연구소의 과제를 수행하기 위하여 외부의 전문가를 공동연구진(연구협력직)으로 선정(연심 결과 첨부)하여 개인과 진행하는 위탁 연구 계약\n• 위탁연구계약 : 각 대학 산학협력단, 출연연 연구기관 등의 외부기관과 연구소 과제의 일부분을 위탁하는 계약\n• 협동연구계약 : 위탁계약과 형식은 유사하나 경제인문사회연구회의 협동연구과제 수행을 위하여 타기관과 체결하는 공동연구 계약\n• 기타 일반계약 : 구매, 인쇄등록업체 선정, 전산정보 유지보수 등의 일반 계약'
      }
    };
    writeJSON('process_rules.json', seedRules);
  }

  // 5. Vendors (거래처)
  const vendors = readJSON('vendors.json', null);
  if (!vendors || vendors.length === 0) {
    const seedVendors = [
      {
        id: 'V-2026-001',
        category: '법인사업자',
        code: 'C-10021',
        name: '(주)한국인공지능시스템즈',
        birthDate: '1978-04-15',
        businessNumber: '123-81-99887',
        corporateNumber: '110111-1234567',
        address: '서울특별시 서초구 강남대로 321',
        detailAddress: '인공지능혁신빌딩 8층 802호',
        bankName: '국민은행',
        accountNumber: '920301-01-234567',
        representative: '홍길동',
        contact: '02-555-0199'
      },
      {
        id: 'V-2026-002',
        category: '법인사업자',
        code: 'C-10022',
        name: '(주)글로벌리서치센터',
        birthDate: '1982-11-20',
        businessNumber: '214-86-12345',
        corporateNumber: '110111-9876543',
        address: '서울특별시 마포구 월드컵북로 402',
        detailAddress: 'KGIT센터 15층',
        bankName: '신한은행',
        accountNumber: '110-345-678901',
        representative: '강대성',
        contact: '02-789-0123'
      },
      {
        id: 'V-2026-003',
        category: '비영리법인',
        code: 'C-10023',
        name: '서울대학교 산학협력단',
        birthDate: '-',
        businessNumber: '119-82-00312',
        corporateNumber: '114571-0001234',
        address: '서울특별시 관악구 관악로 1',
        detailAddress: '산학협력본부 3층',
        bankName: '우리은행',
        accountNumber: '1002-123-456789',
        representative: '유홍림',
        contact: '02-880-1114'
      },
      {
        id: 'V-2026-004',
        category: '개인사업자',
        code: 'C-10024',
        name: '디지털인쇄디자인',
        birthDate: '1980-09-05',
        businessNumber: '105-23-45678',
        corporateNumber: '-',
        address: '서울특별시 중구 을지로 100',
        detailAddress: '파인에비뉴 2층',
        bankName: '하나은행',
        accountNumber: '234-910001-12345',
        representative: '이영희',
        contact: '02-2266-7788'
      },
      {
        id: 'V-2026-005',
        category: '연구협력직(개인)',
        code: 'C-10025',
        name: '최민수 박사',
        birthDate: '1975-03-22',
        businessNumber: '-',
        corporateNumber: '-',
        address: '경기도 성남시 분당구 판교역로 235',
        detailAddress: '에이치스퀘어 N동',
        bankName: '카카오뱅크',
        accountNumber: '3333-01-2345678',
        representative: '최민수',
        contact: '010-9876-5432'
      }
    ];
    writeJSON('vendors.json', seedVendors);
  }
};

initData();

// Helper: Calculate Due Status
const calculateDeadlines = (contract) => {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let inspectionStatus = {
    required: true,
    dueDate: '',
    daysRemaining: null,
    isOverdue: false,
    overdueDays: 0,
    statusText: '정상',
    badgeClass: 'badge-info'
  };

  let paymentStatus = {
    required: false,
    dueDate: '',
    daysRemaining: null,
    isOverdue: false,
    overdueDays: 0,
    statusText: '해당없음',
    badgeClass: 'badge-secondary'
  };

  // 1. 검수보고서 14일 계산
  if (contract.endDate) {
    const end = new Date(contract.endDate);
    end.setHours(0, 0, 0, 0);
    const inspectDue = new Date(end);
    inspectDue.setDate(inspectDue.getDate() + 14);
    
    inspectionStatus.dueDate = inspectDue.toISOString().split('T')[0];

    if (contract.inspectionSubmitted && contract.inspectionDate) {
      inspectionStatus.statusText = `검수완료 (${contract.inspectionDate})`;
      inspectionStatus.badgeClass = 'badge-success';
    } else {
      const diffTime = inspectDue.getTime() - today.getTime();
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      inspectionStatus.daysRemaining = diffDays;

      if (diffDays < 0) {
        inspectionStatus.isOverdue = true;
        inspectionStatus.overdueDays = Math.abs(diffDays);
        inspectionStatus.statusText = `⚠️ 14일 초과 (+${Math.abs(diffDays)}일)`;
        inspectionStatus.badgeClass = 'badge-danger font-weight-bold animate-pulse';
      } else if (diffDays <= 3) {
        inspectionStatus.statusText = `⏰ 마감임박 (D-${diffDays})`;
        inspectionStatus.badgeClass = 'badge-warning';
      } else {
        inspectionStatus.statusText = `진행중 (D-${diffDays})`;
        inspectionStatus.badgeClass = 'badge-primary';
      }
    }
  }

  // 2. 대금지출 5일 계산
  if (contract.paymentDraftDate) {
    paymentStatus.required = true;
    const draft = new Date(contract.paymentDraftDate);
    draft.setHours(0, 0, 0, 0);
    const payDue = new Date(draft);
    payDue.setDate(payDue.getDate() + 5);

    paymentStatus.dueDate = payDue.toISOString().split('T')[0];

    if (contract.isPaid && contract.paymentPaidDate) {
      paymentStatus.statusText = `지출완료 (${contract.paymentPaidDate})`;
      paymentStatus.badgeClass = 'badge-success';
    } else {
      const diffTime = payDue.getTime() - today.getTime();
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      paymentStatus.daysRemaining = diffDays;

      if (diffDays < 0) {
        paymentStatus.isOverdue = true;
        paymentStatus.overdueDays = Math.abs(diffDays);
        paymentStatus.statusText = `⚠️ 5일 초과 (+${Math.abs(diffDays)}일)`;
        paymentStatus.badgeClass = 'badge-danger font-weight-bold animate-pulse';
      } else if (diffDays <= 2) {
        paymentStatus.statusText = `⏰ 마감임박 (D-${diffDays})`;
        paymentStatus.badgeClass = 'badge-warning';
      } else {
        paymentStatus.statusText = `대기중 (D-${diffDays})`;
        paymentStatus.badgeClass = 'badge-primary';
      }
    }
  }

  return { inspectionStatus, paymentStatus };
};

// ================= AUTH API =================
app.post('/api/auth/login', (req, res) => {
  const { username, password } = req.body;
  // Standard admin credential check
  if (username === 'admin' && (password === 'admin123' || password === 'admin' || password === '1234')) {
    res.json({
      success: true,
      user: {
        username: 'admin',
        name: '시스템 관리자',
        role: 'ADMIN'
      },
      token: 'admin-auth-token-' + Date.now()
    });
  } else {
    res.status(401).json({ success: false, message: '아이디 또는 비밀번호가 일치하지 않습니다. (기본 관리자 계정: admin / admin123)' });
  }
});

// ================= PROCESS GUIDE API =================
app.get('/api/process-guide', (req, res) => {
  const guide = readJSON('process_guide.json', []);
  res.json(guide);
});

app.put('/api/process-guide', (req, res) => {
  const updatedGuide = req.body;
  if (!Array.isArray(updatedGuide)) {
    return res.status(400).json({ error: 'Process guide must be an array' });
  }
  writeJSON('process_guide.json', updatedGuide);
  res.json({ success: true, guide: updatedGuide });
});

// Process Classification Rules API (가, 나, 다)
app.get('/api/process-rules', (req, res) => {
  const rules = readJSON('process_rules.json', {});
  res.json(rules);
});

app.put('/api/process-rules', (req, res) => {
  const updatedRules = req.body;
  writeJSON('process_rules.json', updatedRules);
  res.json({ success: true, rules: updatedRules });
});

// Template Document Upload for Process Guide Step
app.post('/api/process-guide/:stepId/templates', uploadTemplate.single('file'), (req, res) => {
  const guide = readJSON('process_guide.json', []);
  const stepIdx = guide.findIndex(s => s.id === req.params.stepId);
  if (stepIdx === -1) return res.status(404).json({ error: 'Step not found' });

  const { name } = req.body;
  const file = req.file;

  if (!guide[stepIdx].templates) guide[stepIdx].templates = [];

  const newTemplate = {
    id: `template-${Date.now()}`,
    name: name || (file ? file.originalname : '표준 서식'),
    fileName: file ? file.originalname : '',
    filePath: file ? `/uploads/templates/${file.filename}` : '',
    fileSize: file ? file.size : 0,
    uploadedAt: new Date().toISOString().split('T')[0]
  };

  guide[stepIdx].templates.push(newTemplate);
  writeJSON('process_guide.json', guide);
  res.status(201).json(guide[stepIdx]);
});

app.delete('/api/process-guide/:stepId/templates/:templateId', (req, res) => {
  const guide = readJSON('process_guide.json', []);
  const stepIdx = guide.findIndex(s => s.id === req.params.stepId);
  if (stepIdx === -1) return res.status(404).json({ error: 'Step not found' });

  guide[stepIdx].templates = (guide[stepIdx].templates || []).filter(t => t.id !== req.params.templateId);
  writeJSON('process_guide.json', guide);
  res.json(guide[stepIdx]);
});

// Projects API
app.get('/api/projects', (req, res) => {
  const projects = readJSON('projects.json', []);
  const { year } = req.query;
  if (year) {
    return res.json(projects.filter(p => String(p.year) === String(year)));
  }
  res.json(projects);
});

app.post('/api/projects', (req, res) => {
  const projects = readJSON('projects.json', []);
  const newProject = {
    id: `P-${req.body.year || new Date().getFullYear()}-${Date.now().toString().slice(-4)}`,
    year: Number(req.body.year) || new Date().getFullYear(),
    code: req.body.code || `PRJ-${req.body.year}-${Date.now().toString().slice(-2)}`,
    name: req.body.name,
    department: req.body.department,
    manager: req.body.manager,
    email: req.body.email,
    budget: Number(req.body.budget) || 0,
    startDate: req.body.startDate,
    endDate: req.body.endDate,
    description: req.body.description || '',
    files: []
  };
  projects.unshift(newProject);
  writeJSON('projects.json', projects);
  res.status(201).json(newProject);
});

app.put('/api/projects/:id', (req, res) => {
  const projects = readJSON('projects.json', []);
  const idx = projects.findIndex(p => p.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: 'Project not found' });
  
  projects[idx] = { ...projects[idx], ...req.body };
  writeJSON('projects.json', projects);
  res.json(projects[idx]);
});

app.delete('/api/projects/:id', (req, res) => {
  let projects = readJSON('projects.json', []);
  projects = projects.filter(p => p.id !== req.params.id);
  writeJSON('projects.json', projects);
  res.json({ success: true, id: req.params.id });
});

// Project File Upload API
app.post('/api/projects/:id/files', upload.single('file'), (req, res) => {
  const projects = readJSON('projects.json', []);
  const project = projects.find(p => p.id === req.params.id);
  if (!project) return res.status(404).json({ error: 'Project not found' });

  const { title } = req.body;
  const file = req.file;
  if (!project.files) project.files = [];

  const newFile = {
    id: `pfile-${Date.now()}`,
    title: title || (file ? file.originalname : '과제 첨부파일'),
    fileName: file ? file.originalname : '',
    filePath: file ? `/uploads/${file.filename}` : '',
    fileSize: file ? file.size : 0,
    uploadedAt: new Date().toISOString().split('T')[0]
  };

  project.files.push(newFile);
  writeJSON('projects.json', projects);
  res.status(201).json(project);
});

app.delete('/api/projects/:id/files/:fileId', (req, res) => {
  const projects = readJSON('projects.json', []);
  const project = projects.find(p => p.id === req.params.id);
  if (!project) return res.status(404).json({ error: 'Project not found' });

  project.files = (project.files || []).filter(f => f.id !== req.params.fileId);
  writeJSON('projects.json', projects);
  res.json(project);
});

// ================= PROJECT EXCEL TEMPLATE & UPLOAD API =================
app.get('/api/projects/excel-template', (req, res) => {
  const wb = XLSX.utils.book_new();
  const sampleData = [
    {
      '당해년도': 2026,
      '과제코드': 'PRJ-2026-01',
      '과제명': 'AI 기반 공공정책 데이터 분석 고도화 연구',
      '주관부서': '디지털혁신연구실',
      '연구책임자': '김철수 책임연구원',
      '책임자이메일': 'chulsoo.kim@re.kr',
      '과제총예산(원)': 150000000,
      '연구시작일': '2026-01-01',
      '연구종료일': '2026-12-31',
      '과제설명': '공공정책 수립을 위한 초거대 AI 기반 분석 모델 개발'
    },
    {
      '당해년도': 2026,
      '과제코드': 'PRJ-2026-02',
      '과제명': '국가 균형발전 지표 조사 및 영향도 분석',
      '주관부서': '지역발전연구본부',
      '연구책임자': '이영희 선임연구원',
      '책임자이메일': 'younghee.lee@re.kr',
      '과제총예산(원)': 80000000,
      '연구시작일': '2026-02-01',
      '연구종료일': '2026-11-30',
      '과제설명': '지역 소멸 위기 대응 균형발전 지표 체계 구축'
    },
    {
      '당해년도': 2026,
      '과제코드': 'PRJ-2026-03',
      '과제명': '2026 경제인문사회연구회 협동연구 : 미래 인구구조 대응 전략',
      '주관부서': '사회정책연구실',
      '연구책임자': '박민수 연구위원',
      '책임자이메일': 'minsu.park@re.kr',
      '과제총예산(원)': 60000000,
      '연구시작일': '2026-03-01',
      '연구종료일': '2026-12-31',
      '과제설명': '초고령사회 대비 정책 연구'
    }
  ];
  // 현재 화면(선택 년도)에 표시된 과제 목록 기준으로 서식 생성 (없으면 샘플 행)
  const year = Number(req.query.year) || null;
  const current = readJSON('projects.json', [])
    .filter(p => !year || Number(p.year) === year)
    .map(p => ({
      '당해년도': Number(p.year) || '',
      '과제코드': p.code || '',
      '과제명': p.name || '',
      '주관부서': p.department || '',
      '연구책임자': p.manager || '',
      '책임자이메일': p.email || '',
      '과제총예산(원)': Number(p.budget) || 0,
      '연구시작일': p.startDate || '',
      '연구종료일': p.endDate || '',
      '과제설명': p.description || ''
    }));
  const rowsData = current.length ? current : sampleData;
  const ws = XLSX.utils.json_to_sheet(rowsData);
  ws['!cols'] = [8, 16, 44, 20, 18, 26, 16, 12, 12, 44].map(wch => ({ wch }));
  XLSX.utils.book_append_sheet(wb, ws, '과제목록');
  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

  res.setHeader('Content-Disposition', `attachment; filename="project_list.xlsx"; filename*=UTF-8''${encodeURIComponent('과제목록_' + (year || '전체') + '.xlsx')}`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buffer);
});

app.post('/api/projects/upload-excel', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: '엑셀 파일이 첨부되지 않았습니다.' });

  try {
    const workbook = XLSX.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rawRows = XLSX.utils.sheet_to_json(sheet);

    if (!rawRows || rawRows.length === 0) {
      return res.status(400).json({ error: '엑셀 파일에 데이터 행이 없습니다.' });
    }

    const projects = readJSON('projects.json', []);
    let addedCount = 0;
    let updatedCount = 0;

    for (const row of rawRows) {
      const year = Number(row['당해년도'] || row['연도'] || row['년도'] || row['year']) || new Date().getFullYear();
      const code = String(row['과제코드'] || row['과제번호'] || row['코드'] || row['code'] || `PRJ-${year}-${Date.now().toString().slice(-3)}`).trim();
      const name = String(row['과제명'] || row['연구과제명'] || row['과제제목'] || row['name'] || '').trim();
      if (!name) continue;

      const department = String(row['주관부서'] || row['부서'] || row['담당부서'] || row['department'] || '연구부서').trim();
      const manager = String(row['연구책임자'] || row['책임자'] || row['담당자'] || row['manager'] || '담당연구원').trim();
      const email = String(row['책임자이메일'] || row['이메일'] || row['email'] || '').trim();
      const budget = Number(row['과제총예산(원)'] || row['과제예산'] || row['예산'] || row['budget'] || 0);
      const startDate = String(row['연구시작일'] || row['시작일'] || row['startDate'] || '').trim();
      const endDate = String(row['연구종료일'] || row['종료일'] || row['endDate'] || '').trim();
      const description = String(row['과제설명'] || row['개요'] || row['비고'] || row['description'] || '').trim();

      const existingIdx = projects.findIndex(p => p.code === code && Number(p.year) === year);
      if (existingIdx !== -1) {
        projects[existingIdx] = {
          ...projects[existingIdx],
          name,
          department,
          manager,
          email,
          budget,
          startDate: startDate || projects[existingIdx].startDate,
          endDate: endDate || projects[existingIdx].endDate,
          description: description || projects[existingIdx].description
        };
        updatedCount++;
      } else {
        projects.push({
          id: `P-${year}-${Date.now().toString().slice(-4)}-${Math.floor(Math.random()*1000)}`,
          year,
          code,
          name,
          department,
          manager,
          email,
          budget,
          startDate,
          endDate,
          description,
          files: []
        });
        addedCount++;
      }
    }

    writeJSON('projects.json', projects);
    res.json({
      success: true,
      message: `총 ${rawRows.length}건 중 신규 과제 ${addedCount}건 등록, 기존 ${updatedCount}건 수정 완료되었습니다.`,
      addedCount,
      updatedCount,
      totalProjects: projects.length
    });
  } catch (err) {
    console.error('Error processing project excel upload:', err);
    res.status(500).json({ error: '엑셀 파일 처리 중 오류가 발생했습니다: ' + err.message });
  } finally {
    try { if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path); } catch (e) {}
  }
});

// ================= VENDORS (거래처 관리) API =================
app.get('/api/vendors', (req, res) => {
  const vendors = readJSON('vendors.json', []);
  const { search } = req.query;
  if (search) {
    const q = search.toLowerCase();
    const filtered = vendors.filter(v =>
      (v.name && v.name.toLowerCase().includes(q)) ||
      (v.code && v.code.toLowerCase().includes(q)) ||
      (v.businessNumber && v.businessNumber.includes(q)) ||
      (v.category && v.category.includes(q)) ||
      (v.bankName && v.bankName.includes(q))
    );
    return res.json(filtered);
  }
  res.json(vendors);
});

app.post('/api/vendors', (req, res) => {
  const vendors = readJSON('vendors.json', []);
  const newVendor = {
    id: `V-${Date.now()}`,
    category: req.body.category || '법인사업자',
    code: req.body.code || `C-${Date.now().toString().slice(-5)}`,
    name: req.body.name,
    birthDate: req.body.birthDate || '-',
    businessNumber: req.body.businessNumber || '-',
    corporateNumber: req.body.corporateNumber || '-',
    address: req.body.address || '',
    detailAddress: req.body.detailAddress || '',
    bankName: req.body.bankName || '',
    accountNumber: req.body.accountNumber || '',
    representative: req.body.representative || '',
    contact: req.body.contact || '',
    createdAt: new Date().toISOString().split('T')[0]
  };
  vendors.unshift(newVendor);
  writeJSON('vendors.json', vendors);
  res.status(201).json(newVendor);
});

app.put('/api/vendors/:id', (req, res) => {
  const vendors = readJSON('vendors.json', []);
  const idx = vendors.findIndex(v => v.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: 'Vendor not found' });

  vendors[idx] = { ...vendors[idx], ...req.body };
  writeJSON('vendors.json', vendors);
  res.json(vendors[idx]);
});

app.delete('/api/vendors/:id', (req, res) => {
  let vendors = readJSON('vendors.json', []);
  vendors = vendors.filter(v => v.id !== req.params.id);
  writeJSON('vendors.json', vendors);
  res.json({ success: true, id: req.params.id });
});

app.get('/api/vendors/excel-template', (req, res) => {
  const wb = XLSX.utils.book_new();
  const sampleData = [
    {
      '거래처구분': '법인사업자',
      '거래처코드': 'C-10021',
      '거래처명': '(주)한국인공지능시스템즈',
      '생년월일': '1978-04-15',
      '사업자번호': '123-81-99887',
      '법인번호': '110111-1234567',
      '주소': '서울특별시 서초구 강남대로 321',
      '상세주소': '인공지능혁신빌딩 8층 802호',
      '은행명': '국민은행',
      '계좌번호': '920301-01-234567'
    },
    {
      '거래처구분': '법인사업자',
      '거래처코드': 'C-10022',
      '거래처명': '(주)글로벌리서치센터',
      '생년월일': '1982-11-20',
      '사업자번호': '214-86-12345',
      '법인번호': '110111-9876543',
      '주소': '서울특별시 마포구 월드컵북로 402',
      '상세주소': 'KGIT센터 15층',
      '은행명': '신한은행',
      '계좌번호': '110-345-678901'
    },
    {
      '거래처구분': '비영리법인',
      '거래처코드': 'C-10023',
      '거래처명': '서울대학교 산학협력단',
      '생년월일': '-',
      '사업자번호': '119-82-00312',
      '법인번호': '114571-0001234',
      '주소': '서울특별시 관악구 관악로 1',
      '상세주소': '산학협력본부 3층',
      '은행명': '우리은행',
      '계좌번호': '1002-123-456789'
    }
  ];
  // 현재 화면(검색어/거래처구분 필터)에 표시된 거래처 목록 기준으로 서식 생성 (없으면 샘플 행)
  const q = String(req.query.search || '').toLowerCase().trim();
  const cat = String(req.query.category || '');
  const current = readJSON('vendors.json', [])
    .filter(v => !cat || v.category === cat)
    .filter(v => !q || [v.name, v.code, v.businessNumber, v.corporateNumber, v.address, v.bankName, v.accountNumber, v.representative]
      .some(f => f && String(f).toLowerCase().includes(q)))
    .map(v => ({
      '거래처구분': v.category || '',
      '거래처코드': v.code || '',
      '거래처명': v.name || '',
      '생년월일': v.birthDate || '-',
      '사업자번호': v.businessNumber || '-',
      '법인번호': v.corporateNumber || '-',
      '주소': v.address || '',
      '상세주소': v.detailAddress || '',
      '은행명': v.bankName || '',
      '계좌번호': v.accountNumber || '',
      '대표자': v.representative || '',
      '연락처': v.contact || ''
    }));
  const rowsData = current.length ? current : sampleData;
  const ws = XLSX.utils.json_to_sheet(rowsData);
  ws['!cols'] = [14, 12, 28, 12, 16, 18, 36, 28, 12, 22, 12, 26].map(wch => ({ wch }));
  XLSX.utils.book_append_sheet(wb, ws, '거래처목록');
  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

  res.setHeader('Content-Disposition', `attachment; filename="vendor_list.xlsx"; filename*=UTF-8''${encodeURIComponent('거래처_목록.xlsx')}`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buffer);
});

app.post('/api/vendors/upload-excel', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: '엑셀 파일이 첨부되지 않았습니다.' });

  try {
    const workbook = XLSX.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rawRows = XLSX.utils.sheet_to_json(sheet);

    if (!rawRows || rawRows.length === 0) {
      return res.status(400).json({ error: '엑셀 파일에 데이터 행이 없습니다.' });
    }

    const vendors = readJSON('vendors.json', []);
    let addedCount = 0;
    let updatedCount = 0;

    for (const row of rawRows) {
      const category = String(row['거래처구분'] || row['구분'] || row['category'] || '법인사업자').trim();
      const code = String(row['거래처코드'] || row['코드'] || row['code'] || `C-${Date.now().toString().slice(-4)}`).trim();
      const name = String(row['거래처명'] || row['업체명'] || row['상호'] || row['name'] || '').trim();
      if (!name) continue;

      const birthDate = String(row['생년월일'] || row['birthDate'] || '-').trim();
      const businessNumber = String(row['사업자번호'] || row['사업자등록번호'] || row['businessNumber'] || '-').trim();
      const corporateNumber = String(row['법인번호'] || row['법인등록번호'] || row['corporateNumber'] || '-').trim();
      const address = String(row['주소'] || row['소재지'] || row['address'] || '').trim();
      const detailAddress = String(row['상세주소'] || row['detailAddress'] || '').trim();
      const bankName = String(row['은행명'] || row['금융기관'] || row['은행'] || row['bankName'] || '').trim();
      const accountNumber = String(row['계좌번호'] || row['accountNumber'] || '').trim();
      const representative = String(row['대표자'] || row['대표자명'] || row['representative'] || '').trim();
      const contact = String(row['연락처'] || row['전화번호'] || row['contact'] || '').trim();

      const existingIdx = vendors.findIndex(v =>
        (v.code && v.code === code) ||
        (businessNumber !== '-' && v.businessNumber === businessNumber) ||
        (v.name === name)
      );

      if (existingIdx !== -1) {
        vendors[existingIdx] = {
          ...vendors[existingIdx],
          category,
          code,
          name,
          birthDate: birthDate !== '-' ? birthDate : vendors[existingIdx].birthDate,
          businessNumber: businessNumber !== '-' ? businessNumber : vendors[existingIdx].businessNumber,
          corporateNumber: corporateNumber !== '-' ? corporateNumber : vendors[existingIdx].corporateNumber,
          address: address || vendors[existingIdx].address,
          detailAddress: detailAddress || vendors[existingIdx].detailAddress,
          bankName: bankName || vendors[existingIdx].bankName,
          accountNumber: accountNumber || vendors[existingIdx].accountNumber,
          representative: representative || vendors[existingIdx].representative,
          contact: contact || vendors[existingIdx].contact
        };
        updatedCount++;
      } else {
        vendors.push({
          id: `V-${Date.now()}-${Math.floor(Math.random()*1000)}`,
          category,
          code,
          name,
          birthDate,
          businessNumber,
          corporateNumber,
          address,
          detailAddress,
          bankName,
          accountNumber,
          representative,
          contact,
          createdAt: new Date().toISOString().split('T')[0]
        });
        addedCount++;
      }
    }

    writeJSON('vendors.json', vendors);
    res.json({
      success: true,
      message: `총 ${rawRows.length}건 중 신규 거래처 ${addedCount}건 등록, 기존 ${updatedCount}건 수정 완료되었습니다.`,
      addedCount,
      updatedCount,
      totalVendors: vendors.length
    });
  } catch (err) {
    console.error('Error processing vendor excel upload:', err);
    res.status(500).json({ error: '거래처 엑셀 파일 처리 중 오류가 발생했습니다: ' + err.message });
  } finally {
    try { if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path); } catch (e) {}
  }
});

// Contract Step Workflow Processing API
app.post('/api/contracts/:id/process-step', (req, res) => {
  const contracts = readJSON('contracts.json', []);
  const idx = contracts.findIndex(c => c.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: 'Contract not found' });

  const contract = contracts[idx];
  const { targetStep, status, dateUpdates, stepNote } = req.body;

  if (targetStep) contract.progressStep = Number(targetStep);
  if (status) contract.status = status;

  if (dateUpdates) {
    if (dateUpdates.startDate) contract.startDate = dateUpdates.startDate;
    if (dateUpdates.endDate) contract.endDate = dateUpdates.endDate;
    if (dateUpdates.inspectionDate) {
      contract.inspectionDate = dateUpdates.inspectionDate;
      contract.inspectionSubmitted = true;
    }
    if (dateUpdates.paymentDraftDate) contract.paymentDraftDate = dateUpdates.paymentDraftDate;
    if (dateUpdates.paymentPaidDate) {
      contract.paymentPaidDate = dateUpdates.paymentPaidDate;
      contract.isPaid = true;
    }
  }

  // 자동 기한 재계산
  if (contract.endDate) {
    contract.inspectionDueDate = new Date(new Date(contract.endDate).getTime() + 14 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
  }
  if (contract.paymentDraftDate) {
    contract.paymentDueDate = new Date(new Date(contract.paymentDraftDate).getTime() + 5 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
  }

  if (!contract.stepHistory) contract.stepHistory = [];
  contract.stepHistory.push({
    step: contract.progressStep,
    status: contract.status,
    note: stepNote || '',
    processedAt: new Date().toISOString()
  });

  contracts[idx] = contract;
  writeJSON('contracts.json', contracts);

  const deadlines = calculateDeadlines(contract);
  res.json({ ...contract, ...deadlines });
});

// Contracts API
app.get('/api/contracts', (req, res) => {
  const contracts = readJSON('contracts.json', []);
  const { year, search, nature, method, status, progressStep, overdueOnly } = req.query;

  let list = contracts.map(c => {
    const deadlines = calculateDeadlines(c);
    return { ...c, ...deadlines };
  });

  if (year) {
    list = list.filter(c => String(c.year) === String(year));
  }
  if (nature) {
    list = list.filter(c => c.contractNature === nature);
  }
  if (method) {
    list = list.filter(c => c.contractType === method);
  }
  if (status) {
    list = list.filter(c => c.status === status);
  }
  if (progressStep) {
    list = list.filter(c => String(c.progressStep) === String(progressStep));
  }
  if (overdueOnly === 'true') {
    list = list.filter(c => c.inspectionStatus.isOverdue || c.paymentStatus.isOverdue);
  }
  if (search) {
    const s = search.toLowerCase();
    list = list.filter(c =>
      c.contractName.toLowerCase().includes(s) ||
      (c.contractorName && c.contractorName.toLowerCase().includes(s)) ||
      (c.projectName && c.projectName.toLowerCase().includes(s)) ||
      (c.manager && c.manager.toLowerCase().includes(s)) ||
      (c.department && c.department.toLowerCase().includes(s))
    );
  }

  res.json(list);
});

app.get('/api/contracts/:id', (req, res) => {
  const contracts = readJSON('contracts.json', []);
  const contract = contracts.find(c => c.id === req.params.id);
  if (!contract) return res.status(404).json({ error: 'Contract not found' });
  
  const deadlines = calculateDeadlines(contract);
  res.json({ ...contract, ...deadlines });
});

app.post('/api/contracts', (req, res) => {
  const contracts = readJSON('contracts.json', []);
  const body = req.body;
  const year = body.year || new Date().getFullYear();
  const newId = `C-${year}-${Date.now().toString().slice(-4)}`;

  // Default 서류 프리셋 생성
  const defaultDocs = [
    { id: `doc-${Date.now()}-1`, category: '계약체결', title: '계약서 원본', fileName: '', uploadedAt: '', isDefault: true, required: true },
    { id: `doc-${Date.now()}-2`, category: '계약체결', title: '보안서약서', fileName: '', uploadedAt: '', isDefault: true, required: true },
    { id: `doc-${Date.now()}-3`, category: '계약체결', title: '계약보증보험증권(지급각서)', fileName: '', uploadedAt: '', isDefault: true, required: true }
  ];

  if (body.contractType === '수의계약') {
    if (Number(body.contractAmount) <= 5000000) {
      defaultDocs.push({ id: `doc-${Date.now()}-4`, category: '계약체결', title: '단일견적서 (500만원 이하)', fileName: '', uploadedAt: '', isDefault: true, required: true });
    } else {
      defaultDocs.push({ id: `doc-${Date.now()}-4`, category: '계약체결', title: '견적서 (주업체)', fileName: '', uploadedAt: '', isDefault: true, required: true });
      defaultDocs.push({ id: `doc-${Date.now()}-5`, category: '계약체결', title: '비교견적서 (타업체 1종 이상)', fileName: '', uploadedAt: '', isDefault: true, required: true });
    }
  } else {
    // 공개경쟁입찰
    defaultDocs.push({ id: `doc-${Date.now()}-4`, category: '입찰공고', title: '제안요청서 및 입찰공고문 (15일 공고)', fileName: '', uploadedAt: '', isDefault: true, required: true });
    defaultDocs.push({ id: `doc-${Date.now()}-5`, category: '기술평가', title: '기술평가회 평가결과표', fileName: '', uploadedAt: '', isDefault: true, required: true });
  }

  if (body.contractNature === '공동연구진계약') {
    defaultDocs.push({ id: `doc-${Date.now()}-6`, category: '계약체결', title: '연구심의(연심) 결과서', fileName: '', uploadedAt: '', isDefault: true, required: true });
  }

  // 선급금 관련 기본서류
  if (body.advancePayment && body.advancePayment.hasAdvance) {
    defaultDocs.push({ id: `doc-${Date.now()}-7`, category: '선급금', title: '선금보증보험증권 (계약종료일+60일)', fileName: '', uploadedAt: '', isDefault: true, required: true });
    defaultDocs.push({ id: `doc-${Date.now()}-8`, category: '선급금', title: '선금사용계획서', fileName: '', uploadedAt: '', isDefault: true, required: true });
  }

  // 준공 및 검수 기본서류
  defaultDocs.push({ id: `doc-${Date.now()}-9`, category: '준공시점', title: '준공계', fileName: '', uploadedAt: '', isDefault: true, required: true });
  defaultDocs.push({ id: `doc-${Date.now()}-10`, category: '준공시점', title: '대금청구서(세금계산서)', fileName: '', uploadedAt: '', isDefault: true, required: true });
  defaultDocs.push({ id: `doc-${Date.now()}-11`, category: '검수시점', title: '검수보고서 (종료일 14일 이내 필수)', fileName: '', uploadedAt: '', isDefault: true, required: true });

  const inspectDueDate = body.endDate ? new Date(new Date(body.endDate).getTime() + 14 * 24 * 60 * 60 * 1000).toISOString().split('T')[0] : '';

  const newContract = {
    id: newId,
    year: Number(year),
    projectId: body.projectId || '',
    projectName: body.projectName || '',
    contractName: body.contractName || '',
    contractType: body.contractType || '수의계약',
    contractNature: body.contractNature || '조사용역계약',
    contractAmount: Number(body.contractAmount) || 0,
    contractorName: body.contractorName || '',
    businessNumber: body.businessNumber || '',
    representative: body.representative || '',
    contractorContact: body.contractorContact || '',
    department: body.department || '',
    manager: body.manager || '',
    managerEmail: body.managerEmail || '',
    startDate: body.startDate || '',
    endDate: body.endDate || '',
    advancePayment: body.advancePayment || { hasAdvance: false },
    progressStep: Number(body.progressStep) || 3,
    status: body.status || '계약체결',
    completionDate: body.completionDate || '',
    inspectionDueDate: inspectDueDate,
    inspectionDate: body.inspectionDate || '',
    inspectionSubmitted: !!body.inspectionDate,
    paymentDraftDate: body.paymentDraftDate || '',
    paymentDueDate: body.paymentDraftDate ? new Date(new Date(body.paymentDraftDate).getTime() + 5 * 24 * 60 * 60 * 1000).toISOString().split('T')[0] : '',
    paymentPaidDate: body.paymentPaidDate || '',
    isPaid: !!body.paymentPaidDate,
    documents: defaultDocs,
    notes: body.notes || ''
  };

  contracts.unshift(newContract);
  writeJSON('contracts.json', contracts);
  res.status(201).json(newContract);
});

app.put('/api/contracts/:id', (req, res) => {
  const contracts = readJSON('contracts.json', []);
  const idx = contracts.findIndex(c => c.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: 'Contract not found' });

  const existing = contracts[idx];
  const updated = { ...existing, ...req.body };

  // 자동 기한 재계산
  if (updated.endDate) {
    updated.inspectionDueDate = new Date(new Date(updated.endDate).getTime() + 14 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
  }
  if (updated.inspectionDate) {
    updated.inspectionSubmitted = true;
  }
  if (updated.paymentDraftDate) {
    updated.paymentDueDate = new Date(new Date(updated.paymentDraftDate).getTime() + 5 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
  }
  if (updated.paymentPaidDate) {
    updated.isPaid = true;
  }

  contracts[idx] = updated;
  writeJSON('contracts.json', contracts);
  res.json(updated);
});

app.delete('/api/contracts/:id', (req, res) => {
  let contracts = readJSON('contracts.json', []);
  contracts = contracts.filter(c => c.id !== req.params.id);
  writeJSON('contracts.json', contracts);
  res.json({ success: true, id: req.params.id });
});

// File Upload for Contract Documents
app.post('/api/contracts/:id/documents', upload.single('file'), (req, res) => {
  const contracts = readJSON('contracts.json', []);
  const contract = contracts.find(c => c.id === req.params.id);
  if (!contract) return res.status(404).json({ error: 'Contract not found' });

  const { docId, title, category, required } = req.body;
  const file = req.file;

  if (!contract.documents) contract.documents = [];

  const existingDocIdx = docId ? contract.documents.findIndex(d => d.id === docId) : -1;

  if (existingDocIdx !== -1) {
    contract.documents[existingDocIdx].title = title || contract.documents[existingDocIdx].title;
    contract.documents[existingDocIdx].category = category || contract.documents[existingDocIdx].category;
    if (file) {
      contract.documents[existingDocIdx].fileName = file.originalname;
      contract.documents[existingDocIdx].filePath = `/uploads/${file.filename}`;
      contract.documents[existingDocIdx].fileSize = file.size;
      contract.documents[existingDocIdx].uploadedAt = new Date().toISOString().split('T')[0];
    }
  } else {
    const newDoc = {
      id: `doc-${Date.now()}`,
      category: category || '추가서류',
      title: title || (file ? file.originalname : '추가 서류'),
      fileName: file ? file.originalname : '',
      filePath: file ? `/uploads/${file.filename}` : '',
      fileSize: file ? file.size : 0,
      uploadedAt: file ? new Date().toISOString().split('T')[0] : '',
      isDefault: false,
      required: required === 'true' || required === true
    };
    contract.documents.push(newDoc);
  }

  if (title && title.includes('검수보고서') && file && !contract.inspectionDate) {
    contract.inspectionDate = new Date().toISOString().split('T')[0];
    contract.inspectionSubmitted = true;
  }

  writeJSON('contracts.json', contracts);
  res.json(contract);
});

// Delete document or file attachment
app.delete('/api/contracts/:id/documents/:docId', (req, res) => {
  const contracts = readJSON('contracts.json', []);
  const contract = contracts.find(c => c.id === req.params.id);
  if (!contract) return res.status(404).json({ error: 'Contract not found' });

  const docIdx = contract.documents.findIndex(d => d.id === req.params.docId);
  if (docIdx === -1) return res.status(404).json({ error: 'Document not found' });

  const doc = contract.documents[docIdx];
  if (doc.isDefault) {
    doc.fileName = '';
    doc.filePath = '';
    doc.uploadedAt = '';
    doc.fileSize = 0;
  } else {
    contract.documents.splice(docIdx, 1);
  }

  writeJSON('contracts.json', contracts);
  res.json(contract);
});

// Mail notification endpoint
app.post('/api/notifications/send', (req, res) => {
  const { contractId, recipient, recipientEmail, subject, content, type } = req.body;
  const notifications = readJSON('notifications.json', []);

  const newLog = {
    id: `NOTI-${Date.now()}`,
    contractId,
    recipient,
    recipientEmail,
    subject,
    content,
    type: type || '검수보고서 제출 알림',
    sentAt: new Date().toISOString(),
    status: '발송완료(시뮬레이션)'
  };

  notifications.unshift(newLog);
  writeJSON('notifications.json', notifications);

  res.json({ success: true, log: newLog, message: `[${recipientEmail}] 으로 알림 메일이 성공적으로 전송되었습니다.` });
});

app.get('/api/notifications', (req, res) => {
  const notifications = readJSON('notifications.json', []);
  res.json(notifications);
});

// Delete single notification log
app.delete('/api/notifications/:id', (req, res) => {
  let notifications = readJSON('notifications.json', []);
  const beforeLen = notifications.length;
  notifications = notifications.filter(n => n.id !== req.params.id);
  if (notifications.length === beforeLen) {
    return res.status(404).json({ error: '발송 이력을 찾을 수 없습니다.' });
  }
  writeJSON('notifications.json', notifications);
  res.json({ success: true, message: '발송 이력이 삭제되었습니다.' });
});

// Clear all notification logs
app.delete('/api/notifications', (req, res) => {
  writeJSON('notifications.json', []);
  res.json({ success: true, message: '모든 발송 이력이 초기화되었습니다.' });
});

// Dashboard Stats
app.get('/api/dashboard/stats', (req, res) => {
  const contracts = readJSON('contracts.json', []);
  const projects = readJSON('projects.json', []);
  const { year } = req.query;

  let cList = contracts;
  let pList = projects;
  if (year) {
    cList = cList.filter(c => String(c.year) === String(year));
    pList = pList.filter(p => String(p.year) === String(year));
  }

  const enriched = cList.map(c => ({ ...c, ...calculateDeadlines(c) }));

  const totalContracts = enriched.length;
  const totalAmount = enriched.reduce((sum, c) => sum + (c.contractAmount || 0), 0);
  const overdueInspectionCount = enriched.filter(c => c.inspectionStatus.isOverdue).length;
  const overduePaymentCount = enriched.filter(c => c.paymentStatus.isOverdue).length;
  const urgentInspectionCount = enriched.filter(c => !c.inspectionStatus.isOverdue && c.inspectionStatus.daysRemaining <= 3 && !c.inspectionSubmitted).length;
  const completedCount = enriched.filter(c => c.isPaid).length;
  const activeCount = enriched.filter(c => !c.isPaid).length;

  const methodStats = {
    '공개경쟁입찰': enriched.filter(c => c.contractType === '공개경쟁입찰').length,
    '수의계약': enriched.filter(c => c.contractType === '수의계약').length
  };

  const natureStats = {
    '조사용역계약': enriched.filter(c => c.contractNature === '조사용역계약').length,
    '공동연구진계약': enriched.filter(c => c.contractNature === '공동연구진계약').length,
    '위탁연구계약': enriched.filter(c => c.contractNature === '위탁연구계약').length,
    '협동연구계약': enriched.filter(c => c.contractNature === '협동연구계약').length,
    '기타일반계약': enriched.filter(c => c.contractNature === '기타일반계약' || c.contractNature === '기타 일반계약').length
  };

  // 진행 상태별 분포
  const statusStats = {
    '계획기안': enriched.filter(c => c.status === '계획기안' || c.progressStep === 1).length,
    '협조문작성': enriched.filter(c => c.status === '협조문작성' || c.progressStep === 2).length,
    '계약체결/공고': enriched.filter(c => c.status === '계약체결' || c.status === '공고/협조' || c.progressStep === 3).length,
    '선급금/착수': enriched.filter(c => c.status === '진행중' || c.progressStep === 4).length,
    '준공검수': enriched.filter(c => c.status === '준공검수' || c.progressStep === 5).length,
    '대금기안/지출': enriched.filter(c => c.status === '대금기안' || c.status === '지출완료' || c.progressStep === 6).length
  };

  res.json({
    year: year || '전체',
    totalProjects: pList.length,
    totalContracts,
    totalAmount,
    overdueInspectionCount,
    overduePaymentCount,
    urgentInspectionCount,
    activeCount,
    completedCount,
    methodStats,
    natureStats,
    statusStats,
    overdueContracts: enriched.filter(c => c.inspectionStatus.isOverdue || c.paymentStatus.isOverdue)
  });
});

// Route for User Portal on Port 3000
app.get('/user', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'user.html'));
});

// Dedicated User Portal App on Port 3001
const userApp = express();
userApp.use(cors());
userApp.use(express.json());
userApp.use(express.urlencoded({ extended: true }));

// Shared API & Uploads
userApp.use('/api', app);
userApp.use('/uploads', express.static(UPLOADS_DIR));
userApp.use(express.static(PUBLIC_DIR));

userApp.get('/', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'user.html'));
});
userApp.get('*', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'user.html'));
});

const ADMIN_PORT = PORT;
const USER_PORT = process.env.USER_PORT || 3001;

app.listen(ADMIN_PORT, () => {
  console.log(`====================================================`);
  console.log(`👑 [관리자 화면 (Admin)] 포트 ${ADMIN_PORT} 실행: http://localhost:${ADMIN_PORT}`);
  console.log(`👤 [일반사용자 화면] 바로가기: http://localhost:${ADMIN_PORT}/user`);
  console.log(`====================================================`);
});

userApp.listen(USER_PORT, () => {
  console.log(`👤 [일반사용자 전용 포털 (User)] 포트 ${USER_PORT} 실행: http://localhost:${USER_PORT}`);
  console.log(`====================================================`);
});
