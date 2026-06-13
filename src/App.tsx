import styled from "styled-components";
import { useState } from "react";
import * as XLSX from "xlsx";

const Container = styled.div`
  min-height: 100vh;
  background: linear-gradient(135deg, #eef4ff 0%, #f8fafc 45%, #ffffff 100%);
  padding: 48px 20px;
  color: #0f172a;

  .inner-container {
    max-width: 960px;
    margin: 0 auto;
  }

  .title-box {
    margin-bottom: 28px;
  }

  h1 {
    font-size: 36px;
    margin: 0 0 8px;
  }

  .description {
    color: #64748b;
    font-size: 15px;
  }

  .card {
    background: rgba(255, 255, 255, 0.88);
    border: 1px solid #e2e8f0;
    border-radius: 22px;
    padding: 24px;
    box-shadow: 0 20px 50px rgba(15, 23, 42, 0.08);
    margin-bottom: 24px;
  }

  .control-box {
    display: flex;
    gap: 12px;
    align-items: center;
    flex-wrap: wrap;
  }

  .upload-btn,
  button {
    height: 44px;
    border: 0;
    border-radius: 12px;
    padding: 0 18px;
    font-weight: 700;
    cursor: pointer;
  }

  .upload-btn {
    display: inline-flex;
    align-items: center;
    background: #0f172a;
    color: white;
  }

  select {
    height: 44px;
    min-width: 120px;
    border: 1px solid #cbd5e1;
    border-radius: 12px;
    padding: 0 14px;
    background: white;
    font-weight: 700;
  }

  button {
    background: #2563eb;
    color: white;
  }

  .export-btn {
    background: #16a34a;
  }

  .section-title {
    font-size: 20px;
    margin: 0 0 16px;
  }

  .name-list {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(130px, 1fr));
    gap: 10px;
    list-style: none;
    padding: 0;
    margin: 0;
  }

  .name-list li {
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 10px 12px;
    font-weight: 600;
  }

  .group-wrapper {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
    gap: 16px;
  }

  .group-item {
    background: white;
    border: 1px solid #e2e8f0;
    border-radius: 18px;
    padding: 18px;
    box-shadow: 0 12px 28px rgba(15, 23, 42, 0.06);
  }

  .group-item h3 {
    margin: 0 0 12px;
    color: #2563eb;
  }

  .group-item ul {
    list-style: none;
    padding: 0;
    margin: 0;
  }

  .group-item li {
    padding: 8px 0;
    border-bottom: 1px solid #f1f5f9;
    font-weight: 600;
  }

  .group-item li:last-child {
    border-bottom: 0;
  }
`;

function App() {
  const [names, setNames] = useState<string[]>([]);
  const [groupCount, setGroupCount] = useState<number>(2);
  const [groups, setGroups] = useState<string[][]>([]);

  const handleExcelUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];

    if (!file) return;

    const reader = new FileReader();

    reader.onload = (event) => {
      const data = event.target?.result;

      if (!data) return;

      const workbook = XLSX.read(data, {
        type: "array",
      });

      const sheetName = workbook.SheetNames[0];

      const worksheet = workbook.Sheets[sheetName];

      const rows = XLSX.utils.sheet_to_json<string[]>(worksheet, {
        header: 1,
      });

      // A열만 추출
      const aColumnNames = rows
        .map((row) => row[0])
        .map((name) => String(name ?? "").trim())
        .filter((name) => name.length > 0);

      // 중복 제거
      const uniqueNames = Array.from(new Set(aColumnNames));

      setNames(uniqueNames);

      setGroups([]);
    };

    reader.readAsArrayBuffer(file);
  };

  const handleGroupFormation = () => {
    if (names.length === 0) return;

    const shuffled = [...names];

    // Fisher-Yates Shuffle
    for (let i = shuffled.length - 1; i > 0; i--) {
      const randomIndex = Math.floor(Math.random() * (i + 1));

      [shuffled[i], shuffled[randomIndex]] = [
        shuffled[randomIndex],
        shuffled[i],
      ];
    }

    const result: string[][] = Array.from({ length: groupCount }, () => []);

    shuffled.forEach((name, index) => {
      result[index % groupCount].push(name);
    });

    setGroups(result);
  };

  const handleExportExcel = () => {
    if (groups.length === 0) return;

    const maxRows = Math.max(...groups.map((group) => group.length));

    const exportData: Record<string, string>[] = [];

    const header: Record<string, string> = {};

    groups.forEach((_, index) => {
      header[`group${index + 1}`] = `${index + 1}조`;
    });

    exportData.push(header);

    for (let rowIndex = 0; rowIndex < maxRows; rowIndex++) {
      const row: Record<string, string> = {};

      groups.forEach((group, groupIndex) => {
        row[`group${groupIndex + 1}`] = group[rowIndex] || "";
      });

      exportData.push(row);
    }

    const worksheet = XLSX.utils.json_to_sheet(exportData, {
      skipHeader: true,
    });

    worksheet["!cols"] = groups.map(() => ({
      wch: 15,
    }));

    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(workbook, worksheet, "조편성 결과");

    XLSX.writeFile(workbook, "조편성_결과.xlsx");
  };

  return (
    <Container>
      <div className="inner-container">
        <div className="title-box">
          <h1>랜덤 조편성</h1>
          <p className="description">
            엑셀 A열 명단을 불러와 원하는 개수의 조로 랜덤 편성합니다.
          </p>
        </div>

        <div className="card">
          <div className="control-box">
            <input
              id="excel-upload"
              type="file"
              accept=".xlsx,.xls"
              onChange={handleExcelUpload}
              hidden
            />

            <label htmlFor="excel-upload" className="upload-btn">
              엑셀 파일 선택
            </label>

            {names.length >= 2 && (
              <select
                value={groupCount}
                onChange={(e) => setGroupCount(Number(e.target.value))}
              >
                {Array.from({ length: names.length - 1 }, (_, index) => (
                  <option key={index + 2} value={index + 2}>
                    {index + 2}개 조
                  </option>
                ))}
              </select>
            )}

            <button onClick={handleGroupFormation}>조편성 하기</button>

            <button className="export-btn" onClick={handleExportExcel}>
              엑셀 내보내기
            </button>
          </div>
        </div>

        <div className="card">
          <h2 className="section-title">불러온 명단 ({names.length}명)</h2>

          <ul className="name-list">
            {names.map((name, index) => (
              <li key={`${name}-${index}`}>
                {index + 1}. {name}
              </li>
            ))}
          </ul>
        </div>

        {groups.length > 0 && (
          <div className="card">
            <h2 className="section-title">조편성 결과</h2>

            <div className="group-wrapper">
              {groups.map((group, index) => (
                <div className="group-item" key={index}>
                  <h3>
                    {index + 1}조 · {group.length}명
                  </h3>

                  <ul>
                    {group.map((member) => (
                      <li key={member}>{member}</li>
                    ))}
                  </ul>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    </Container>
  );
}

export default App;
