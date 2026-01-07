import { NextRequest, NextResponse } from 'next/server';
import puppeteer from 'puppeteer';

export async function POST(request: NextRequest) {
  try {
    console.log('=== Bernard Referral PDF API 호출됨 ===');
    const data = await request.json();
    console.log('=== PDF 생성 요청 ===');
    console.log('전체 데이터:', JSON.stringify(data, null, 2));
    console.log('환자명:', data.patientName);
    console.log('첨부 파일 개수:', data.uploadedFiles ? data.uploadedFiles.length : 0);
    
    // 필수 필드 검증
    if (!data.patientName) {
      console.error('❌ 필수 필드 누락: patientName');
      return NextResponse.json({ error: 'patientName is required' }, { status: 400 });
    }
    if (!data.office) {
      console.error('❌ 필수 필드 누락: office');
      return NextResponse.json({ error: 'office is required' }, { status: 400 });
    }
    if (!data.type) {
      console.error('❌ 필수 필드 누락: type');
      return NextResponse.json({ error: 'type is required' }, { status: 400 });
    }
    
    if (data.uploadedFiles && data.uploadedFiles.length > 0) {
      data.uploadedFiles.forEach((file, index) => {
        console.log(`파일 ${index + 1}: ${file.name} (${file.size} bytes, 데이터: ${!!file.data})`);
      });
    }
    

    // 첨부 파일 데이터 처리 함수
    const processFileData = async (file) => {
      console.log('파일 처리 시작:', file.name);
      console.log('파일 타입:', file.type);
      console.log('파일 데이터 상태:', {
        hasData: !!file.data,
        dataType: typeof file.data,
        dataPreview: file.data ? file.data.substring(0, 100) + '...' : 'No data'
      });
      
      // PDF 파일은 더 이상 지원하지 않음
      if (file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf')) {
        console.log('PDF 파일 감지 - 지원하지 않는 파일 형식:', file.name);
        file.isPdf = false;
        file.isImage = false;
        file.data = null;
        return file;
      }
      
      // 이미지 파일 처리
      if (!file.data) {
        console.log('파일 데이터 없음:', file.name);
        return file;
      }
      
      // Firebase Storage URL인 경우
      if (file.data.includes('firebasestorage.googleapis.com')) {
        console.log('Firebase Storage URL 감지:', file.name);
        
        // 이미지 파일인지 확인
        const isImageFile = /\.(jpg|jpeg|png|gif|bmp|webp)$/i.test(file.name);
        
        if (isImageFile) {
          try {
            console.log('이미지 파일 Base64 변환 시도:', file.name);
            // Firebase Storage URL에서 토큰 제거하여 공개 URL로 변환
            const url = new URL(file.data);
            url.searchParams.delete('token');
            const publicUrl = url.toString();
            
            console.log('공개 URL로 변환:', publicUrl);
            
            // 이미지 다운로드 시도
            const response = await fetch(publicUrl, {
              headers: {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
              }
            });
            
            if (response.ok) {
              const arrayBuffer = await response.arrayBuffer();
              const base64 = Buffer.from(arrayBuffer).toString('base64');
              const mimeType = response.headers.get('content-type') || 'image/jpeg';
              file.data = `data:${mimeType};base64,${base64}`;
              console.log('이미지를 Base64로 변환 완료:', file.name, `크기: ${base64.length} bytes`);
            } else {
              console.log('Firebase Storage URL 접근 실패:', file.name, response.status, response.statusText);
              // URL을 그대로 사용
              file.data = publicUrl;
            }
          } catch (error) {
            console.log('Firebase Storage URL 처리 오류:', file.name, error.message);
            // URL을 그대로 사용
            const url = new URL(file.data);
            url.searchParams.delete('token');
            file.data = url.toString();
          }
        } else {
          // 이미지가 아닌 파일은 URL만 정리
          console.log('이미지가 아닌 파일, URL만 정리:', file.name);
          const url = new URL(file.data);
          url.searchParams.delete('token');
          file.data = url.toString();
        }
      }
      
      console.log('파일 처리 완료:', file.name);
      return file;
    };

    // 첨부 파일 처리
    if (data.uploadedFiles && data.uploadedFiles.length > 0) {
      console.log('첨부 파일 처리 시작...');
      data.uploadedFiles = await Promise.all(data.uploadedFiles.map(processFileData));
      console.log('첨부 파일 처리 완료');
    }

    // HTML 생성
    const htmlContent = `
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>Bernard Referral Form</title>
        <style>
            body {
                font-family: Arial, sans-serif;
                margin: 0;
                padding: 10px;
                background: white;
                line-height: 1.4;
                font-size: 12px;
                color: #000;
            }
            .form-section {
                background: white;
                padding: 0;
                margin-bottom: 10px;
                border: none;
                box-shadow: none;
            }
            .form-section h3 {
                color: #000;
                border-bottom: 1px solid #000;
                padding-bottom: 8px;
                margin-bottom: 20px;
                font-size: 20px;
                font-weight: bold;
                text-transform: uppercase;
                text-align: center;
            }
            .form-row {
                display: flex;
                margin-bottom: 8px;
                gap: 15px;
            }
            .form-group {
                flex: 1;
            }
            .info-table {
                width: 100%;
                border-collapse: collapse;
                margin-bottom: 20px;
            }
            .label-cell {
                font-weight: bold;
                color: #000;
                font-size: 12px;
                text-transform: uppercase;
                padding: 8px 12px;
                border: 1px solid #000;
                background: #f8f9fa;
                width: 30%;
                vertical-align: top;
            }
            .value-cell {
                padding: 8px 12px;
                border: 1px solid #000;
                font-size: 12px;
                color: #000;
                width: 70%;
                vertical-align: top;
            }
            .form-group-single {
                width: 100%;
                margin-bottom: 4px;
            }
            .form-group label {
                display: block;
                font-weight: bold;
                margin-bottom: 3px;
                color: #000;
                font-size: 12px;
                text-transform: uppercase;
            }
            .form-group div {
                padding: 5px 8px;
                background: white;
                border: 1px solid #000;
                border-radius: 0;
                min-height: 18px;
                font-size: 12px;
                color: #000;
            }
            .full-width {
                width: 100%;
            }
            .page-break {
                page-break-before: always;
            }
            @page {
                size: A4;
                margin: 10mm;
            }
            @media print {
                body {
                    -webkit-print-color-adjust: exact;
                    print-color-adjust: exact;
                }
                .no-break {
                    page-break-inside: avoid;
                }
            }
            .pdf-container {
                width: 100%;
                height: 100vh;
                display: flex;
                flex-direction: column;
                align-items: center;
                justify-content: center;
                background: white;
                padding: 0;
                box-sizing: border-box;
            }
            .pdf-iframe {
                width: 100%;
                height: 80vh;
                border: none;
                border-radius: 0;
                box-shadow: none;
                transform: scale(1);
                object-fit: contain;
            }
            @media print {
                .page-break {
                    page-break-before: always;
                }
            }
        </style>
    </head>
    <body>
        <div class="form-section">
            <h3>Referral</h3>
            <table class="info-table">
                <tr>
                    <td class="label-cell">Date:</td>
                    <td class="value-cell">${data.date}</td>
                </tr>
                <tr>
                    <td class="label-cell">Office:</td>
                    <td class="value-cell">${data.office}</td>
                </tr>
                <tr>
                    <td class="label-cell">Type:</td>
                    <td class="value-cell">${data.type}</td>
                </tr>
                <tr>
                    <td class="label-cell">Patient Name:</td>
                    <td class="value-cell">${data.patientName}</td>
                </tr>
                <tr>
                    <td class="label-cell">Date of Birth:</td>
                    <td class="value-cell">${data.dob}</td>
                </tr>
                <tr>
                    <td class="label-cell">Insurance:</td>
                    <td class="value-cell">${data.insurance} ${data.insuranceOther ? `(${data.insuranceOther})` : ''}</td>
                </tr>
                <tr>
                    <td class="label-cell">Behavior:</td>
                    <td class="value-cell">${data.behavior}</td>
                </tr>
                <tr>
                    <td class="label-cell">Medical Condition:</td>
                    <td class="value-cell">${data.medicalCondition} ${data.medicalConditionDetails ? `(${data.medicalConditionDetails})` : ''}</td>
                </tr>
                <tr>
                    <td class="label-cell">Selected Teeth Numbers:</td>
                    <td class="value-cell">${data.selectedNumbers}</td>
                </tr>
                <tr>
                    <td class="label-cell">Remarks:</td>
                    <td class="value-cell">${data.remarks}</td>
                </tr>
            </table>
        </div>

        ${data.uploadedFiles && data.uploadedFiles.length > 0 ? `
        <!-- 첨부 파일 섹션 -->
        ${data.uploadedFiles.map((file, index) => {
          console.log(`HTML 생성 중 - 파일 ${index + 1}:`, {
            name: file.name,
            hasData: !!file.data,
            dataLength: file.data ? file.data.length : 0
          });
          
          const isImage = file.type && file.type.startsWith('image/');
          const isPdf = file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf');
          
          console.log(`파일 ${file.name} - isImage: ${isImage}, isPdf: ${isPdf}, hasData: ${!!file.data}`);
          
          // 페이지 브레이크 추가 (첫 번째 파일이 아닌 경우)
          const pageBreak = index > 0 ? 'page-break-before: always;' : '';
          
          if (isImage && file.data) {
            console.log(`이미지 파일 처리: ${file.name}`);
            return `
              <div style="width: 100%; height: calc(100vh - 20mm); ${pageBreak} display: flex; align-items: center; justify-content: center; background: white; padding: 0; margin: 0; box-sizing: border-box; page-break-after: avoid;">
                <img src="${file.data}" style="max-width: 100%; max-height: calc(100vh - 40mm); object-fit: contain;" alt="${file.name}" />
              </div>
            `;
          } else {
            // 기타 파일 또는 데이터가 없는 경우
            const fileIcon = isPdf ? '📄' : (isImage ? '🖼️' : '📎');
            const fileSize = file.size ? (file.size / 1024 / 1024).toFixed(1) + 'MB' : 'Unknown size';
            
            return `
              <div style="width: 100%; height: calc(100vh - 20mm); ${pageBreak} display: flex; flex-direction: column; align-items: center; justify-content: center; background: #f8f9fa; padding: 20px; box-sizing: border-box; page-break-after: avoid;">
                <div style="text-align: center; margin-bottom: 20px;">
                  <h3 style="color: #333; font-size: 24px; margin-bottom: 10px;">Attached File</h3>
                  <p style="color: #666; font-size: 16px; margin: 0;">${file.name}</p>
                </div>
                <div style="flex: 1; display: flex; align-items: center; justify-content: center; width: 100%;">
                  <div style="width: 400px; height: 500px; background: linear-gradient(135deg, #6c757d, #495057); color: white; display: flex; flex-direction: column; align-items: center; justify-content: center; border-radius: 12px; box-shadow: 0 8px 16px rgba(0,0,0,0.1); position: relative; overflow: hidden;">
                    <div style="position: absolute; top: 20px; right: 20px; font-size: 40px; opacity: 0.3;">${fileIcon}</div>
                    <div style="position: absolute; bottom: 20px; left: 20px; font-size: 20px; opacity: 0.3;">FILE</div>
                    <div style="text-align: center; z-index: 1;">
                      <div style="font-size: 80px; margin-bottom: 20px;">${fileIcon}</div>
                      <div style="font-size: 24px; font-weight: bold; margin-bottom: 10px;">File Attached</div>
                      <div style="font-size: 16px; text-align: center; word-break: break-word; padding: 0 20px; margin-bottom: 20px; line-height: 1.2;">${file.name}</div>
                      <div style="font-size: 14px; opacity: 0.8; margin-bottom: 20px;">Size: ${fileSize}</div>
                      ${file.data && file.data.startsWith('http') ? 
                        `<a href="${file.data}" target="_blank" style="background: rgba(255,255,255,0.25); color: white; padding: 12px 24px; border-radius: 6px; text-decoration: none; font-size: 16px; border: 1px solid rgba(255,255,255,0.4); display: inline-block; transition: all 0.2s;">📥 Download</a>` : 
                        `<div style="font-size: 14px; opacity: 0.7; background: rgba(255,255,255,0.1); padding: 8px 16px; border-radius: 6px;">File attached</div>`
                      }
                    </div>
                  </div>
                </div>
                <div style="text-align: center; margin-top: 20px; color: #666; font-size: 14px;">
                  Size: ${fileSize}
                </div>
              </div>
            `;
          }
        }).join('')}
        ` : ''}

    </body>
    </html>
    `;
    
    // Puppeteer 브라우저 초기화
    const browser = await puppeteer.launch({
      headless: true,
      args: ['--no-sandbox', '--disable-setuid-sandbox']
    });
    
    const page = await browser.newPage();
    await page.setContent(htmlContent);
    
    // 이미지 로딩을 위한 대기 시간 추가
    await new Promise(resolve => setTimeout(resolve, 3000));
    
    // 페이지 높이 확인 및 조정
    const bodyHeight = await page.evaluate(() => {
      return document.body.scrollHeight;
    });
    
    console.log('페이지 전체 높이:', bodyHeight);
    
    // 빈 페이지 방지를 위한 추가 스크립트 실행
    await page.evaluate(() => {
      // 모든 요소의 불필요한 여백 제거
      const elements = document.querySelectorAll('*');
      elements.forEach(el => {
        const style = window.getComputedStyle(el);
        if (style.marginBottom === '0px' && style.paddingBottom === '0px') {
          el.style.marginBottom = '0';
          el.style.paddingBottom = '0';
        }
      });
      
      // 마지막 요소의 여백 제거
      const lastElement = document.body.lastElementChild;
      if (lastElement) {
        lastElement.style.marginBottom = '0';
        lastElement.style.paddingBottom = '0';
        lastElement.style.pageBreakAfter = 'avoid';
      }
    });
    
    // 모든 이미지가 로드될 때까지 대기
    await page.evaluate(() => {
      return new Promise((resolve) => {
        const images = document.querySelectorAll('img');
        let loadedCount = 0;
        const totalImages = images.length;
        
        if (totalImages === 0) {
          resolve();
          return;
        }
        
        images.forEach((img) => {
          if (img.complete) {
            loadedCount++;
          } else {
            img.onload = () => {
              loadedCount++;
              if (loadedCount === totalImages) {
                resolve();
              }
            };
            img.onerror = () => {
              loadedCount++;
              if (loadedCount === totalImages) {
                resolve();
              }
            };
          }
        });
        
        if (loadedCount === totalImages) {
          resolve();
        }
      });
    });
    
    const pdf = await page.pdf({
      format: 'A4',
      printBackground: true,
      margin: {
        top: '10mm',
        right: '10mm',
        bottom: '10mm',
        left: '10mm'
      },
      preferCSSPageSize: false,
      displayHeaderFooter: false,
      omitBackground: false,
      tagged: false,
      scale: 1.0
    });
    
    await browser.close();
    
    // 파일명에서 특수문자 제거 및 ASCII 문자만 사용
    const sanitizedPatientName = data.patientName.replace(/[^a-zA-Z0-9_-]/g, '_');
    const sanitizedDate = data.date.replace(/[^a-zA-Z0-9_-]/g, '_');
    const filename = `Bernard_Referral_${sanitizedPatientName}_${sanitizedDate}.pdf`;
    
    return new NextResponse(pdf, {
      headers: {
        'Content-Type': 'application/pdf',
        'Content-Disposition': `attachment; filename="${filename}"`
      }
    });
    
  } catch (error) {
    console.error('❌ PDF 생성 중 오류:', error);
    console.error('❌ 에러 스택:', error.stack);
    return NextResponse.json(
      { 
        error: 'PDF 생성에 실패했습니다.',
        details: error.message,
        stack: error.stack
      },
      { status: 500 }
    );
  }
}