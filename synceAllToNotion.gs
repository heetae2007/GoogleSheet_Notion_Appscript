const NOTION_DATABASE_ID = "#####"; //노션 DB_ID 저장

// ============================================================
// 2. Google Sheet → Notion 전체 동기화
// ============================================================
function syncAllToNotion() {

  const NOTION_TOKEN = getNotionToken();

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const SKIP_TABS = [
    "설정",
    "공통",
    "Sheet1",
    "요약"
  ];

  // ----------------------------------------------------------
  // Google Sheet 티켓 수집
  // key = 티켓번호
  // ----------------------------------------------------------
  const ticketMap = new Map();

  const sheets = ss.getSheets();

  for (const sheet of sheets) {

    const sheetName = sheet.getName();

    if (SKIP_TABS.includes(sheetName)) {
      continue;
    }

    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();

    // 실제 데이터 없음
    if (lastRow < 3 || lastColumn === 0) {
      continue;
    }

    // ========================================================
    // 2행 = 헤더
    // ========================================================
    const headers = sheet
      .getRange(2, 1, 1, lastColumn)
      .getDisplayValues()[0];

    const normalizedHeaders =
      headers.map(normalizeHeader);

    const receivedAtIndex =
      normalizedHeaders.indexOf("*접수일시");

    const ticketNoIndex =
      normalizedHeaders.indexOf("티켓번호");

    const hospitalIndex =
      normalizedHeaders.indexOf("요청병원");

    const partIndex =
      normalizedHeaders.indexOf("관련파트");

    const summaryIndex =
      normalizedHeaders.indexOf("요청내용요약");

    const typeIndex =
      normalizedHeaders.indexOf("*구분");

    // 개발담당자/
    // PIS담당자 대응
    const managerIndex =
      normalizedHeaders.findIndex(header =>
        header.includes("개발담당자") &&
        header.includes("PIS담당자")
      );

    // --------------------------------------------------------
    // 필수 컬럼 체크
    // --------------------------------------------------------
    if (
      receivedAtIndex === -1 ||
      ticketNoIndex === -1 ||
      hospitalIndex === -1 ||
      partIndex === -1 ||
      summaryIndex === -1 ||
      managerIndex === -1 ||
      typeIndex === -1
    ) {

      Logger.log(
        "필수 컬럼 없음 - 건너뜀: " +
        sheetName
      );

      continue;
    }

    // ========================================================
    // 3행부터 실제 데이터
    // ========================================================
    const rows = sheet
      .getRange(
        3,
        1,
        lastRow - 2,
        lastColumn
      )
      .getDisplayValues();

    for (const row of rows) {

      const receivedAtText =
        row[receivedAtIndex];

      if (!receivedAtText) {
        continue;
      }

      const receivedAt =
       parseSheetDate(receivedAtText);

      if (!receivedAt) {
        Logger.log(
          "날짜 변환 실패: " +
          sheetName +
          " / " +
          receivedAtText
        );

        continue;
      }


      // ========================================================
      // 오늘 접수된 티켓만 처리
      // ========================================================
      if (!isToday(receivedAt)) {
        continue;
      }

      function isToday(date) {

        const timeZone = "Asia/Seoul";

        const today =
          Utilities.formatDate(
            new Date(),
            timeZone,
            "yyyy-MM-dd"
          );

        const targetDate =
          Utilities.formatDate(
            date,
            timeZone,
            "yyyy-MM-dd"
          );

        return today === targetDate;
      }

      const ticketNo =
        String(row[ticketNoIndex]).trim();

      if (!ticketNo) {
        continue;
      }

      const ticket = {
        ticketNo: ticketNo,

        summary:
          String(
            row[summaryIndex] || ""
          ).trim(),

        hospital:
          String(
            row[hospitalIndex] || ""
          ).trim(),

        part:
          String(
            row[partIndex] || ""
          ).trim(),

        manager:
          String(
            row[managerIndex] || ""
          ).trim(),

        ticketType:
          String(
            row[typeIndex] || ""
          ).trim(),

        receivedAt:
          receivedAtText,

        sourceSheet:
          sheetName
      };

      // 같은 티켓이 여러 담당자 탭에 존재할 수 있으므로
      // 티켓번호 기준으로 한 번만 저장
      if (!ticketMap.has(ticketNo)) {
        ticketMap.set(ticketNo, ticket);
      }
    }
  }

  Logger.log(
    "동기화 대상 티켓 수: " +
    ticketMap.size
  );


  // ==========================================================
  // 3. Notion에 기존 티켓 조회
  // ==========================================================
  const notionTicketMap =
    getExistingNotionTickets(NOTION_TOKEN);

  Logger.log(
    "Notion 기존 티켓 수: " +
    notionTicketMap.size
  );


  // ==========================================================
  // 4. 신규 / 수정 처리
  // ==========================================================
  let createdCount = 0;
  let updatedCount = 0;
  let errorCount = 0;

  for (const [
    ticketNo,
    ticket
  ] of ticketMap.entries()) {

    try {

      const existingPage =
        notionTicketMap.get(ticketNo);

      if (existingPage) {

        // Google Sheet의 * 접수일시
        const receivedAt =
          parseSheetDate(ticket.receivedAt);

        const now =
          new Date();

        const TWO_HOURS =
          2 * 60 * 60 * 1000;

        const diffMs =
          now.getTime() - receivedAt.getTime();


        // 접수일시로부터 2시간 이내인 경우만 업데이트
        if (
          diffMs >= 0 &&
          diffMs <= TWO_HOURS
        ) {

          updateNotionTicket(
            NOTION_TOKEN,
            existingPage.pageId,
            ticket
          );

          updatedCount++;

          Logger.log(
            "UPDATE (접수 후 2시간 이내): " +
            ticketNo +
            " / " +
            ticket.summary
          );

        } else {

          Logger.log(
            "SKIP UPDATE (접수 후 2시간 초과): " +
            ticketNo +
            " / 접수일시: " +
            ticket.receivedAt
          );
        }

      } else {

        // 신규 티켓은 시간과 관계없이 생성
        const newPageId =
          createNotionTicket(
            NOTION_TOKEN,
            ticket
          );

        notionTicketMap.set(
          ticketNo,
          {
            pageId: newPageId
          }
        );

        createdCount++;

        Logger.log(
          "CREATE: " +
          ticketNo +
          " / " +
          ticket.summary
        );
      }

      // Notion API 호출 속도 조절
      Utilities.sleep(350);

    } catch (e) {

      errorCount++;

      Logger.log(
        "ERROR: " +
        ticketNo +
        " / " +
        e.message
      );
    }
  }


  Logger.log("================================");
  Logger.log("Notion 동기화 완료");
  Logger.log("신규 생성: " + createdCount);
  Logger.log("업데이트: " + updatedCount);
  Logger.log("오류: " + errorCount);
  Logger.log("================================");
}


// ============================================================
// Notion 기존 티켓 전체 조회
// ============================================================
function getExistingNotionTickets(
  notionToken
) {

  const result = new Map();

  let hasMore = true;
  let startCursor = null;

  while (hasMore) {

    const url =
      "https://api.notion.com/v1/databases/" +
      NOTION_DATABASE_ID +
      "/query";

    const today =
      Utilities.formatDate(
        new Date(),
        "Asia/Seoul",
        "yyyy-MM-dd"
      );

    const body = {
      page_size: 100,

      filter: {
        property: "게시일",
        date: {
          equals: today
        }
      }
    };

    if (startCursor) {
      body.start_cursor = startCursor;
    }

    const response =
      UrlFetchApp.fetch(
        url,
        {
          method: "post",

          headers: notionHeaders(
            notionToken
          ),

          payload:
            JSON.stringify(body),

          muteHttpExceptions: true
        }
      );

    const status =
      response.getResponseCode();

    const text =
      response.getContentText();

    if (status !== 200) {
      throw new Error(
        "Notion 조회 실패 (" +
        status +
        "): " +
        text
      );
    }

    const data =
      JSON.parse(text);

    for (const page of data.results) {

      const property =
        page.properties["티켓번호"];

      if (
        property &&
        property.type === "number" &&
        property.number !== null
      ) {

        const ticketNo =
          String(property.number);

        result.set(
          ticketNo,
          {
            pageId: page.id
          }
        );
      }
    }

    hasMore =
      data.has_more === true;

    startCursor =
      data.next_cursor || null;
  }

  return result;
}


// ============================================================
// 신규 Notion 페이지 생성
// ============================================================
function createNotionTicket(
  notionToken,
  ticket
) {

  const payload = {
    parent: {
      database_id:
        NOTION_DATABASE_ID
    },

    properties:
      buildNotionProperties(ticket),

    // 신규 티켓에만 템플릿 생성
    children:
      getTemplateBlocks(
        ticket.ticketType
      )
  };

  const response =
    UrlFetchApp.fetch(
      "https://api.notion.com/v1/pages",
      {
        method: "post",

        headers:
          notionHeaders(
            notionToken
          ),

        payload:
          JSON.stringify(payload),

        muteHttpExceptions: true
      }
    );

  const status =
    response.getResponseCode();

  const text =
    response.getContentText();

  if (status !== 200) {
    throw new Error(
      "Notion 생성 실패 (" +
      status +
      "): " +
      text
    );
  }

  const data =
    JSON.parse(text);

  return data.id;
}


// ============================================================
// 기존 Notion 페이지 업데이트
// ============================================================
function updateNotionTicket(
  notionToken,
  pageId,
  ticket
) {

  const payload = {
    properties:
      buildNotionProperties(ticket)
  };

  const response =
    UrlFetchApp.fetch(
      "https://api.notion.com/v1/pages/" +
      pageId,
      {
        method: "patch",

        headers:
          notionHeaders(
            notionToken
          ),

        payload:
          JSON.stringify(payload),

        muteHttpExceptions: true
      }
    );

  const status =
    response.getResponseCode();

  const text =
    response.getContentText();

  if (status !== 200) {
    throw new Error(
      "Notion 수정 실패 (" +
      status +
      "): " +
      text
    );
  }
}


// ============================================================
// Notion 속성 생성
// ============================================================
function buildNotionProperties(
  ticket
) {

  const title =
    ticket.ticketNo +
    ": " +
    (
      ticket.summary ||
      "제목 없음"
    );

  return {

    // --------------------------------------------------------
    // 제목
    // 예:
    // 96255: 퇴원간호확인 취소가 안되어 문의
    // --------------------------------------------------------
    "콘텐츠 이름": {
      title: [
        {
          type: "text",

          text: {
            content:
              truncateText(
                title,
                1900
              )
          }
        }
      ]
    },


    // --------------------------------------------------------
    // Zendesk 티켓번호
    // --------------------------------------------------------
    "티켓번호": {
      number:
        Number(
          ticket.ticketNo
        )
    },


    // --------------------------------------------------------
    // 요청병원
    // --------------------------------------------------------
    "요청병원": {
      rich_text:
        makeRichText(
          ticket.hospital
        )
    },


    // --------------------------------------------------------
    // 관련파트
    // 원무, 진료 등
    // --------------------------------------------------------
    "관련파트": {
      multi_select:
        parseMultiSelect(
          ticket.part
        )
    },


    // --------------------------------------------------------
    // 게시일
    // --------------------------------------------------------
    "게시일": {
      date:
        ticket.receivedAt
          ? {
              start:
                convertToNotionDate(
                  ticket.receivedAt
                )
            }
          : null
    },


    // --------------------------------------------------------
    // 담당자
    // "강지영, 김구현, 김종오"
    // 그대로 텍스트 저장
    // --------------------------------------------------------
    "담당자": {
      rich_text:
        makeRichText(
          ticket.manager
        )
    },


    // --------------------------------------------------------
    // 구분
    // 데이터수정 / 사용법문의 / 기타문의 / 데이터요청 / 오류
    // --------------------------------------------------------
    "구분": {
      multi_select: ticket.ticketType
        ? [
            {
              name: ticket.ticketType
            }
          ]
        : []
    }
  };
}


// ============================================================
// 구분별 Notion 페이지 본문 템플릿
// ============================================================
function getTemplateBlocks(
  ticketType
) {

  switch (
    String(ticketType).trim()
  ) {

    // ========================================================
    // 데이터 수정
    // ========================================================
    case "데이터수정":
      return [
        headingBlock(
          "데이터 수정 내용"
        ),

        paragraphBlock(
          "수정쿼리"
        ),

        codeBlock(
          "",
          "sql"
        )

      ];


    // ========================================================
    // 사용법 문의
    // ========================================================
    case "사용법문의":
      return [
        headingBlock(
          "문의 내용"
        ),

        paragraphBlock(""),

        headingBlock(
          "안내 내용"
        ),

        paragraphBlock("")
      ];


    // ========================================================
    // 기타 문의
    // ========================================================
    case "기타문의":
      return [
        headingBlock(
          "문의 내용"
        ),

        paragraphBlock(""),

        headingBlock(
          "처리 내용"
        ),

        paragraphBlock("")
      ];


    // ========================================================
    // 데이터 요청
    // ========================================================
    case "데이터요청":
      return [
        headingBlock(
          "요청 데이터"
        ),

        paragraphBlock(""),

        headingBlock(
          "추출 쿼리"
        ),
        codeBlock(
          "",
          "sql"
        ),

        paragraphBlock(""),

        headingBlock(
          "전달 내용"
        ),

        paragraphBlock("")
      ];


    // ========================================================
    // 오류
    // ========================================================
    case "오류":
      return [
        headingBlock(
          "오류 현상"
        ),

        paragraphBlock(""),

        headingBlock(
          "재현 방법"
        ),

        paragraphBlock(""),

        headingBlock(
          "원인"
        ),

        paragraphBlock(""),

        headingBlock(
          "조치 내용"
        ),

        paragraphBlock("")
      ];


    // ========================================================
    // 정의되지 않은 구분
    // ========================================================
    default:
      return [
        headingBlock(
          "문의 내용"
        ),

        paragraphBlock(""),

        headingBlock(
          "처리 내용"
        ),

        paragraphBlock("")
      ];
  }
}


// ============================================================
// Heading 블록 생성
// ============================================================
function headingBlock(text) {

  return {
    object: "block",

    type: "heading_2",

    heading_2: {
      rich_text: [
        {
          type: "text",

          text: {
            content: text
          }
        }
      ]
    }
  };
}


// ============================================================
// Paragraph 블록 생성
// ============================================================
function paragraphBlock(text) {

  return {
    object: "block",

    type: "paragraph",

    paragraph: {
      rich_text:
        text
          ? [
              {
                type: "text",

                text: {
                  content: text
                }
              }
            ]
          : []
    }
  };
}
// 코드블럭 추가
function codeBlock(text, language) {
  return {
    object: "block",
    type: "code",
    code: {
      rich_text: text
        ? [
            {
              type: "text",
              text: {
                content: text
              }
            }
          ]
        : [],
      language: language || "plain text"
    }
  };
}
// ============================================================
// 헤더 정규화
// ============================================================
function normalizeHeader(value) {

  return String(value)
    .replace(/\s+/g, "")
    .replace(/"/g, "")
    .trim();
}


// ============================================================
// Google Sheet 날짜 변환
// ============================================================
function parseSheetDate(value) {

  if (!value) {
    return null;
  }

  if (
    value instanceof Date &&
    !isNaN(value.getTime())
  ) {
    return value;
  }

  const text =
    String(value).trim();


  // ==========================================================
  // 예:
  // 9월 29일 9:02
  // 10월 1일 10:05
  // 10월 01일 14:30:00
  // ==========================================================
  let match =
    text.match(
      /^(\d{1,2})월\s*(\d{1,2})일(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/
    );

  if (match) {

    const year = 2026;

    return new Date(
      year,
      Number(match[1]) - 1,
      Number(match[2]),
      Number(match[3] || 0),
      Number(match[4] || 0),
      Number(match[5] || 0)
    );
  }


  // ==========================================================
  // 예:
  // 2026-10-01 10:05
  // 2026.10.01 10:05
  // 2026/10/01 10:05
  // ==========================================================
  match =
    text.match(
      /^(\d{4})[-./](\d{1,2})[-./](\d{1,2})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/
    );

  if (match) {

    return new Date(
      Number(match[1]),
      Number(match[2]) - 1,
      Number(match[3]),
      Number(match[4] || 0),
      Number(match[5] || 0),
      Number(match[6] || 0)
    );
  }

  return null;
}


// ============================================================
// Notion 날짜 문자열 변환
// ============================================================
function convertToNotionDate(value) {

  const date =
    parseSheetDate(value);

  if (!date) {

    throw new Error(
      "날짜 변환 실패: " +
      value
    );
  }

  return Utilities.formatDate(
    date,
    "Asia/Seoul",
    "yyyy-MM-dd'T'HH:mm:ssXXX"
  );
}


// ============================================================
// 관련파트 처리
//
// 예:
// "원무, 진료"
// → 원무 / 진료
// ============================================================
function parseMultiSelect(value) {

  if (!value) {
    return [];
  }

  return String(value)
    .split(",")
    .map(v => v.trim())
    .filter(v => v)
    .map(v => ({
      name: v
    }));
}


// ============================================================
// Notion Rich Text 생성
// ============================================================
function makeRichText(value) {

  if (!value) {
    return [];
  }

  return [
    {
      type: "text",

      text: {
        content:
          truncateText(
            String(value),
            1900
          )
      }
    }
  ];
}


// ============================================================
// 긴 문자열 방지
// ============================================================
function truncateText(
  value,
  maxLength
) {

  const text =
    String(value || "");

  if (
    text.length <= maxLength
  ) {
    return text;
  }

  return text.substring(
    0,
    maxLength
  );
}


// ============================================================
// Notion Header
// ============================================================
function notionHeaders(
  notionToken
) {

  return {
    "Authorization":
      "Bearer " +
      notionToken,

    "Notion-Version":
      "2022-06-28",

    "Content-Type":
      "application/json"
  };
}


// ============================================================
// Notion Token 가져오기
// ============================================================
function getNotionToken() {

  const token =
    PropertiesService
      .getScriptProperties()
      .getProperty(
        "NOTION_TOKEN"
      );

  if (!token) {

    throw new Error(
      "NOTION_TOKEN이 설정되어 있지 않습니다."
    );
  }

  return token;
}


// ============================================================
// 최초 1회 실행용
//
// 아래 토큰 부분 수정 후 한 번 실행하고,
// 성공하면 함수 안의 실제 토큰은 삭제해도 됨
// ============================================================
function setNotionToken() {

  const token =
    "ntn_#### 최초 1회 노션 토큰 저장";

  PropertiesService
    .getScriptProperties()
    .setProperty(
      "NOTION_TOKEN",
      token
    );

  Logger.log(
    "NOTION_TOKEN 저장 완료"
  );
}
