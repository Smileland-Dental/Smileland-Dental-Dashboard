import { NextRequest, NextResponse } from 'next/server';
import puppeteer from 'puppeteer';

export async function POST(request: NextRequest) {
  try {
    const { contractDate, formData, signatureData, termsSignatureData } = await request.json();

    if (!contractDate || !formData) {
      return NextResponse.json({ success: false, error: 'Contract date and form data are required' });
    }

    // HTML 템플릿 생성
    const html = `
      <!DOCTYPE html>
      <html lang="en">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Orthodontic_Contract_${formData.patientName?.replace(/\s+/g, '_')}_${contractDate}</title>
        <style>
          @media print {
            body { margin: 0.5in; }
          }
          
          * {
            box-sizing: border-box;
          }
          
          body {
            font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
            margin: 0;
            padding: 18px;
            background: #ffffff;
            color: #1a1a1a;
            line-height: 1.3;
            font-size: 9px;
          }
          
          /* Header Section - Readable */
          .header {
            text-align: center;
            padding: 12px 0;
            margin-bottom: 12px;
            background: #f5f5f5;
            border-top: 2px solid #333;
            border-bottom: 2px solid #333;
          }
          
          .header h1 {
            margin: 0 0 6px 0;
            font-size: 16px;
            font-weight: 700;
            color: #1a1a1a;
            letter-spacing: 1.2px;
            text-transform: uppercase;
          }
          
          .header p {
            margin: 3px 0;
            font-size: 8px;
            color: #4a4a4a;
            font-weight: 400;
          }
          
          /* Patient Information - Readable */
          .contract-info {
            margin-bottom: 10px;
            background: #fafafa;
            padding: 8px;
            border: 1px solid #d0d0d0;
          }
          
          .info-row {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 6px;
            margin-bottom: 6px;
            page-break-inside: avoid;
          }
          
          .info-row:last-child {
            margin-bottom: 0;
          }
          
          .info-item {
            background: white;
            padding: 5px 6px;
            border: 1px solid #d0d0d0;
            border-left: 1.5px solid #666;
          }
          
          .info-label {
            font-weight: 600;
            font-size: 7px;
            color: #666;
            margin-bottom: 2px;
            text-transform: uppercase;
            letter-spacing: 0.2px;
          }
          
          .info-value {
            font-size: 9px;
            color: #1a1a1a;
            font-weight: 500;
            min-height: 12px;
          }
          
          /* Treatment Section - Readable */
          .treatment-section {
            margin: 10px 0;
            padding: 8px 10px;
            background: #f9f9f9;
            border: 1px solid #d0d0d0;
            page-break-inside: avoid;
          }
          
          .treatment-label {
            font-weight: 700;
            font-size: 9px;
            margin-bottom: 6px;
            color: #1a1a1a;
            letter-spacing: 0.3px;
            text-transform: uppercase;
          }
          
          .treatment-options {
            display: flex;
            gap: 12px;
            margin-top: 4px;
          }
          
          .treatment-option {
            font-size: 8px;
            padding: 4px 8px;
            background: white;
            border: 1px solid #d0d0d0;
          }
          
          .selected {
            font-weight: 700;
            color: #1a1a1a;
            font-size: 9px;
            background: #e8e8e8;
            border: 1.5px solid #666;
          }
          
          .checkbox {
            display: inline-block;
            width: 12px;
            height: 12px;
            border: 1px solid #666;
            margin-right: 5px;
            vertical-align: middle;
            position: relative;
            background: white;
          }
          
          .checkbox.checked {
            background: #333;
          }
          
          .checkbox.checked::after {
            content: "✓";
            position: absolute;
            top: -2px;
            left: 2px;
            font-size: 10px;
            font-weight: bold;
            color: white;
          }
          
          /* Services Section - Readable */
          .services-section {
            margin: 10px 0;
            page-break-inside: avoid;
          }
          
          .section-header {
            font-weight: 700;
            font-size: 9px;
            color: #1a1a1a;
            margin-bottom: 6px;
            padding: 5px 8px;
            background: #f5f5f5;
            border-left: 2px solid #666;
            text-transform: uppercase;
            letter-spacing: 0.3px;
          }
          
          .services-list {
            list-style: none;
            padding: 0;
            margin: 4px 0;
          }
          
          .services-list li {
            padding: 5px 8px;
            margin: 3px 0;
            background: #fafafa;
            border: 1px solid #d0d0d0;
            display: flex;
            justify-content: space-between;
            align-items: center;
          }
          
          .service-name {
            font-size: 8px;
            color: #1a1a1a;
            font-weight: 500;
          }
          
          .service-price {
            font-weight: 700;
            color: #333;
            font-size: 8px;
          }
          
          /* Total Section - Readable */
          .total-section {
            margin-top: 10px;
            padding: 8px 12px;
            background: #333;
            display: flex;
            justify-content: space-between;
            align-items: center;
            border: 1.5px solid #1a1a1a;
          }
          
          .total-label {
            font-size: 10px;
            font-weight: 700;
            color: white;
            letter-spacing: 0.4px;
            text-transform: uppercase;
          }
          
          .total-amount {
            font-size: 14px;
            font-weight: 700;
            color: white;
          }
          
          /* Footer - Compact */
          .footer {
            margin-top: 20px;
            text-align: center;
            font-size: 8px;
            color: #666;
            border-top: 2px solid #333;
            padding-top: 8px;
            background: #f5f5f5;
          }
          
          .footer p {
            margin: 2px 0;
          }
        </style>
      </head>
      <body>
        <div class="header">
          <h1>ORTHODONTIC CONTRACT</h1>
          <p><strong>Smileland Dental</strong></p>
        </div>
        
        <div class="contract-info">
          <div class="info-row">
            <div class="info-item">
              <div class="info-label">Patient's Name</div>
              <div class="info-value">${formData.patientName || ''}</div>
            </div>
            <div class="info-item">
              <div class="info-label">Date of Birth</div>
              <div class="info-value">${formData.dob || ''}</div>
            </div>
          </div>
          
          <div class="info-row">
            <div class="info-item">
              <div class="info-label">Responsible Party</div>
              <div class="info-value">${formData.responsibleParty || ''}</div>
            </div>
            <div class="info-item">
              <div class="info-label">Relationship</div>
              <div class="info-value">${formData.relationship || ''}</div>
            </div>
          </div>
          
          <div class="info-row">
            <div class="info-item">
              <div class="info-label">Social Security Number</div>
              <div class="info-value">${formData.ssn || ''}</div>
            </div>
            <div class="info-item">
              <div class="info-label">Driver's License</div>
              <div class="info-value">${formData.driversLicense || ''}</div>
            </div>
          </div>
        </div>
        
        <div class="treatment-section">
          <div class="treatment-label">TYPE OF TREATMENT</div>
          <div class="treatment-options">
            <div class="treatment-option">
              <span class="checkbox ${formData.typeOfTreatment === 'Comprehensive' ? 'checked' : ''}"></span>
              <span class="${formData.typeOfTreatment === 'Comprehensive' ? 'selected' : ''}">Comprehensive</span>
            </div>
            <div class="treatment-option">
              <span class="checkbox ${formData.typeOfTreatment === 'Limited' ? 'checked' : ''}"></span>
              <span class="${formData.typeOfTreatment === 'Limited' ? 'selected' : ''}">Limited</span>
            </div>
            <div class="treatment-option">
              <span class="checkbox ${formData.typeOfTreatment === 'Phase I' ? 'checked' : ''}"></span>
              <span class="${formData.typeOfTreatment === 'Phase I' ? 'selected' : ''}">Phase I</span>
            </div>
            <div class="treatment-option">
              <span class="checkbox ${formData.typeOfTreatment === 'Phase II' ? 'checked' : ''}"></span>
              <span class="${formData.typeOfTreatment === 'Phase II' ? 'selected' : ''}">Phase II</span>
            </div>
          </div>
        </div>
        
        ${formData.servicesRequired && formData.servicesRequired.length > 0 ? `
        <div class="services-section">
          <div class="section-header">Services Required</div>
          <ul class="services-list">
            ${formData.servicesRequired.map((service: string) => {
              return `<li>
                <span class="service-name">✓ ${service}</span>
              </li>`;
            }).join('')}
          </ul>
        </div>
        ` : ''}
        
        ${formData.additionalAppliances && formData.additionalAppliances.length > 0 ? `
        <div class="services-section">
          <div class="section-header">Additional Appliance (if necessary)</div>
          <ul class="services-list">
            ${formData.additionalAppliances.map((appliance: string) => {
              const [name, price] = appliance.split('|');
              return `<li>
                <span class="service-name">✓ ${name}</span>
                <span class="service-price">${price}</span>
              </li>`;
            }).join('')}
          </ul>
        </div>
        ` : ''}

        <!-- Payment Options -->
        <div style="margin-top: 12px; display: grid; grid-template-columns: 1fr 1fr; gap: 8px; page-break-inside: avoid;">
          <!-- First Option -->
          <div style="border: 1px solid #666; padding: 8px; background: #fafafa;">
            <h3 style="margin: 0 0 8px 0; text-align: center; color: #1a1a1a; font-size: 9px; font-weight: 700; letter-spacing: 0.4px; padding-bottom: 5px; border-bottom: 1.5px solid #333; text-transform: uppercase;">First Option</h3>
            
            <table style="width: 100%; border-collapse: collapse; font-size: 8px;">
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; width: 55%; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Treatment:</td>
                <td style="padding: 3px 4px; color: #333; font-weight: 600; background: white; border: 1px solid #d0d0d0; border-left: none;">${formData.firstOption?.treatment ? '$' + formData.firstOption.treatment : ''}</td>
              </tr>
              <tr style="height: 2px;"></tr>
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Appliance:</td>
                <td style="padding: 3px 4px; color: #333; font-weight: 600; background: white; border: 1px solid #d0d0d0; border-left: none;">${formData.firstOption?.appliance ? '$' + formData.firstOption.appliance : ''}</td>
              </tr>
              <tr style="height: 2px;"></tr>
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Deposit:</td>
                <td style="padding: 3px 4px; color: #333; font-weight: 600; background: white; border: 1px solid #d0d0d0; border-left: none;">${formData.firstOption?.deposit ? '$' + formData.firstOption.deposit : ''}</td>
              </tr>
              <tr style="height: 2px;"></tr>
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Subtotal:</td>
                <td style="padding: 3px 4px; color: #333; font-weight: 600; background: white; border: 1px solid #d0d0d0; border-left: none;">${formData.firstOption?.subtotal ? '$' + formData.firstOption.subtotal : ''}</td>
              </tr>
              <tr style="height: 2px;"></tr>
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Est. Insurance:</td>
                <td style="padding: 3px 4px; color: #333; font-weight: 600; background: white; border: 1px solid #d0d0d0; border-left: none;">${formData.firstOption?.estimatedInsurance ? '$' + formData.firstOption.estimatedInsurance : ''}</td>
              </tr>
              <tr style="height: 2px;"></tr>
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Net Balance:</td>
                <td style="padding: 3px 4px; color: #333; font-weight: 600; background: white; border: 1px solid #d0d0d0; border-left: none;">${formData.firstOption?.netBalance ? '$' + formData.firstOption.netBalance : ''}</td>
              </tr>
              <tr style="height: 2px;"></tr>
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Est. Trtmt Period:</td>
                <td style="padding: 3px 4px; color: #333; font-weight: 600; background: white; border: 1px solid #d0d0d0; border-left: none;">${formData.firstOption?.estimatedTreatmentPeriod ? '$' + formData.firstOption.estimatedTreatmentPeriod + ' mos' : ''}</td>
              </tr>
              <tr style="height: 2px;"></tr>
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Monthly Payment:</td>
                <td style="padding: 3px 4px; color: #333; font-weight: 600; background: white; border: 1px solid #d0d0d0; border-left: none;">${formData.firstOption?.monthlyPayment ? '$' + formData.firstOption.monthlyPayment : ''}</td>
              </tr>
            </table>
          </div>

          <!-- Second Option -->
          <div style="border: 1px solid #666; padding: 8px; background: #fafafa;">
            <h3 style="margin: 0 0 8px 0; text-align: center; color: #1a1a1a; font-size: 9px; font-weight: 700; letter-spacing: 0.4px; padding-bottom: 5px; border-bottom: 1.5px solid #333; text-transform: uppercase;">Second Option</h3>
            
            <table style="width: 100%; border-collapse: collapse; font-size: 8px;">
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; width: 55%; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Treatment:</td>
                <td style="padding: 3px 4px; color: #333; font-weight: 600; background: white; border: 1px solid #d0d0d0; border-left: none;">${formData.secondOption?.treatment ? '$' + formData.secondOption.treatment : ''}</td>
              </tr>
              <tr style="height: 2px;"></tr>
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Appliance:</td>
                <td style="padding: 3px 4px; color: #333; font-weight: 600; background: white; border: 1px solid #d0d0d0; border-left: none;">${formData.secondOption?.appliance ? '$' + formData.secondOption.appliance : ''}</td>
              </tr>
              <tr style="height: 2px;"></tr>
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Deposit:</td>
                <td style="padding: 3px 4px; color: #333; font-weight: 600; background: white; border: 1px solid #d0d0d0; border-left: none;">${formData.secondOption?.deposit ? '$' + formData.secondOption.deposit : ''}</td>
              </tr>
              <tr style="height: 2px;"></tr>
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Subtotal:</td>
                <td style="padding: 3px 4px; color: #333; font-weight: 600; background: white; border: 1px solid #d0d0d0; border-left: none;">${formData.secondOption?.subtotal ? '$' + formData.secondOption.subtotal : ''}</td>
              </tr>
              <tr style="height: 2px;"></tr>
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Est. Insurance:</td>
                <td style="padding: 3px 4px; color: #333; font-weight: 600; background: white; border: 1px solid #d0d0d0; border-left: none;">${formData.secondOption?.estimatedInsurance ? '$' + formData.secondOption.estimatedInsurance : ''}</td>
              </tr>
              <tr style="height: 2px;"></tr>
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Net Balance:</td>
                <td style="padding: 3px 4px; color: #333; font-weight: 600; background: white; border: 1px solid #d0d0d0; border-left: none;">${formData.secondOption?.netBalance ? '$' + formData.secondOption.netBalance : ''}</td>
              </tr>
              <tr style="height: 2px;"></tr>
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Est. Trtmt Period:</td>
                <td style="padding: 3px 4px; color: #333; font-weight: 600; background: white; border: 1px solid #d0d0d0; border-left: none;">${formData.secondOption?.estimatedTreatmentPeriod ? '$' + formData.secondOption.estimatedTreatmentPeriod + ' mos' : ''}</td>
              </tr>
              <tr style="height: 2px;"></tr>
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Monthly Payment:</td>
                <td style="padding: 3px 4px; color: #333; font-weight: 600; background: white; border: 1px solid #d0d0d0; border-left: none;">${formData.secondOption?.monthlyPayment ? '$' + formData.secondOption.monthlyPayment : ''}</td>
              </tr>
            </table>
          </div>
        </div>

        <!-- Terms and Conditions - Summary -->
        <div style="margin-top: 12px; padding: 8px 10px; background: #f9f9f9; border: 1px solid #d0d0d0; page-break-inside: avoid;">
          <div style="margin: 0; font-size: 7px; line-height: 1.4; color: #4a4a4a;">
            <div style="margin-bottom: 3px; padding-left: 8px; border-left: 2px solid #666;">* The charges for the procedures already performed on patient are not refundable.</div>
            <div style="margin-bottom: 3px; padding-left: 8px; border-left: 2px solid #666;">* We offer an 5% fee reduction for paying in full with cash at the beginning of orthodontic treatment and a 3% fee reduction if paid in full with credit card (Visa or MasterCard).</div>
            <div style="margin-bottom: 3px; padding-left: 8px; border-left: 2px solid #666;">* Clear brackets are available for an additional cost of $300 per arch.</div>
            <div style="margin-bottom: 3px; padding-left: 8px; border-left: 2px solid #666;">* Estimated insurance amount is not a guaranteed amount by no means and subject to change due to variable calculation formulas used by insurance carriers.</div>
            <div style="margin-bottom: 3px; padding-left: 8px; border-left: 2px solid #666;">* An administration fee of $100 will be charged should the account require a new contract.</div>
            <div style="margin-bottom: 0; padding-left: 8px; border-left: 2px solid #666;">* This quote is valid for 30 days.</div>
          </div>
        </div>
        
        <!-- Signature Section -->
        <div style="margin-top: 12px; padding: 10px; background: #fafafa; border: 1px solid #d0d0d0; page-break-inside: avoid;">
          <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 8px; margin-bottom: 8px;">
            <div>
              <div style="margin-bottom: 3px; font-weight: 700; color: #666; font-size: 7px; text-transform: uppercase; letter-spacing: 0.2px;">Quote Presented by:</div>
              <div style="border-bottom: 1px solid #666; padding: 5px 4px; min-height: 14px; background: white; font-weight: 500; color: #1a1a1a; font-size: 8px;">
                ${formData.quotePresentedBy || ''}
              </div>
            </div>
            <div>
              <div style="margin-bottom: 3px; font-weight: 700; color: #666; font-size: 7px; text-transform: uppercase; letter-spacing: 0.2px;">Date:</div>
              <div style="border-bottom: 1px solid #666; padding: 5px 4px; min-height: 14px; background: white; font-weight: 500; color: #1a1a1a; font-size: 8px;">
                ${formData.quotePresentedDate || ''}
              </div>
            </div>
          </div>
          
          <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 8px;">
            <div>
              <div style="margin-bottom: 3px; font-weight: 700; color: #666; font-size: 7px; text-transform: uppercase; letter-spacing: 0.2px;">Signature of Responsible Party:</div>
              ${signatureData ? `
                <div style="border: 1px solid #666; padding: 5px; min-height: 60px; background: white;">
                  <img src="${signatureData}" style="width: 100%; height: 52px; object-fit: contain;" />
                </div>
              ` : `
                <div style="border: 1px dashed #999; padding-top: 60px; background: white;"></div>
              `}
            </div>
            <div>
              <div style="margin-bottom: 3px; font-weight: 700; color: #666; font-size: 7px; text-transform: uppercase; letter-spacing: 0.2px;">Date:</div>
              <div style="border-bottom: 1px solid #666; padding: 5px 4px; min-height: 14px; background: white; font-weight: 500; color: #1a1a1a; font-size: 8px;">
                ${formData.signatureDate || ''}
              </div>
            </div>
          </div>
        </div>
        
        <!-- Detailed Terms and Conditions -->
        <div style="margin-top: 10px; padding: 25px; background-color: #f8f9fa; border-radius: 8px; border: 1px solid #dee2e6;">
          
          <!-- Section 1: Payment Terms -->
          <div style="margin-bottom: 20px; font-size: 11px; line-height: 1.6; color: #495057;">
            <p style="text-align: justify; margin-bottom: 12px;">
              The unpaid balance of $ <span style="border-bottom: 1px solid #333; display: inline-block; min-width: 60px; text-align: center; padding: 0 5px;">${formData.unpaidBalance || '_______'}</span> will be paid in <span style="border-bottom: 1px solid #333; display: inline-block; min-width: 40px; text-align: center; padding: 0 5px;">${formData.paymentMonths || '_____'}</span> months in the amount of $ <span style="border-bottom: 1px solid #333; display: inline-block; min-width: 60px; text-align: center; padding: 0 5px;">${formData.monthlyPaymentAmount || '_______'}</span>, each due on the 1st or 15th day of the month, beginning on <span style="border-bottom: 1px solid #333; display: inline-block; min-width: 80px; text-align: center; padding: 0 5px;">${formData.paymentBeginDate || '__________'}</span> and continued until the above balance is paid in full. $20 late fee will be applied to your account if your payment is not received within five days of the payment due date. We accept cash, credit card (Visa or MasterCard) or Care Credit as payment at the office location. Personal check can be accepted by mail only.
            </p>
            <p style="text-align: justify; margin-bottom: 12px;">
              The unpaid balance has been divided into monthly payments only to assist you in the payment of the total balance, and has no correlation with your appointments. The total fee must be paid in full prior to removing the orthodontic appliances and placement of retainers. If the account becomes delinquent, patient will not be seen for any treatment except emergency palliative treatment only to relieve pain with a $50 charge. If an account is in default more than three months, the case will be terminated and referred to a collection agency and an additional debanding fee of $600 to be charged to the account.
            </p>
            <div style="margin-top: 8px;">
              <strong>Initial:</strong> <span style="border-bottom: 2px solid #333; display: inline-block; min-width: 80px; padding-left: 5px;">${formData.initial1 || ''}</span>
            </div>
          </div>

          <!-- Section 2: Insurance -->
          <div style="margin-bottom: 20px; font-size: 11px; line-height: 1.6; color: #495057;">
            <p style="text-align: justify; margin-bottom: 8px;">
              If you have orthodontic insurance coverage, please understand that your insurance coverage is a contract between you and your insurance company. We will only assist you in billing your insurance company. The estimated insurance amount is a quote and is subject to change due to variable calculation formulas used by insurance carriers. If your insurance does not pay the estimated amount, you are personally responsible for the total cost of treatment. All scheduled monthly orthodontic visits are mandatory due to the requirements of most insurance carriers. If you miss any of your appointments you will be responsible for the amount which was not paid by your insurance carrier as well as your monthly payment.
            </p>
            <p style="text-align: justify; margin-bottom: 8px;">
              Although after billing the dental insurance, the insurance may send payments to the subscriber (to you or the primary insured) therefore it is your responsibility to bring the payment to the office to clear your outstanding balance. If this payment is sent under the subscriber's name, you may need to cash the check first than bring in the amount to our office. All amounts paid to the subscriber and not paid toward the patient's account will be considered delinquent and therefore can be referred to a collection agency.
            </p>
            <p style="text-align: justify; margin-bottom: 12px;">
              If the patient is under 18, the parents, custodial parent or guardian will be legally responsible for any outstanding balance.
            </p>
            <div style="margin-top: 8px;">
              <strong>Initial:</strong> <span style="border-bottom: 2px solid #333; display: inline-block; min-width: 80px; padding-left: 5px;">${formData.initial2 || ''}</span>
            </div>
          </div>

          <!-- Section 3: General Dentistry -->
          <div style="margin-bottom: 20px; font-size: 11px; line-height: 1.6; color: #495057;">
            <p style="text-align: justify; margin-bottom: 12px;">
              The Orthodontic Treatment Fee does not include any general dentistry such as fillings, expose & bond, extractions, crowns, or dental cleaning. Regular dental check-up and all necessary dental work must be completed and documented by your dentist prior to orthodontic bonding. Please ask your dentist to consult with us should you require any extensive dental treatment. Good oral hygiene is imperative to the success of orthodontic treatment. Should the patient require additional treatment period due to lack of patient's cooperation, such as no elastic wear / missed appointment an additional fee of $100 per month will be charged.
            </p>
            <div style="margin-top: 8px;">
              <strong>Initial:</strong> <span style="border-bottom: 2px solid #333; display: inline-block; min-width: 80px; padding-left: 5px;">${formData.initial3 || ''}</span>
            </div>
          </div>

          <!-- Section 4: Broken Appliances -->
          <div style="margin-bottom: 20px; font-size: 11px; line-height: 1.6; color: #495057;">
            <p style="text-align: justify; margin-bottom: 12px;">
              Replacement of lost or broken appliances will result in appropriate additional charge plus laboratory fee. Broken band metal brackets are $35 and clear brackets are $45 each.
            </p>
            <div style="margin-top: 8px;">
              <strong>Initial:</strong> <span style="border-bottom: 2px solid #333; display: inline-block; min-width: 80px; padding-left: 5px;">${formData.initial4 || ''}</span>
            </div>
          </div>

          <!-- Section 5: Missed Appointments -->
          <div style="margin-bottom: 20px; font-size: 11px; line-height: 1.6; color: #495057;">
            <p style="text-align: justify; margin-bottom: 8px;">
              We will try our best to accommodate you when making appointments. <strong>Any missed months or cancellation without a 24 hour notice is subject to a charge of $20 per appointment</strong>; therefore it is important that you notify us immediately if you need to change your appointment. You will be responsible to contact our office and reschedule any missed appointments. If you arrive more than 15 minutes past your appointment time, you may be asked to reschedule your appointment. Due to heavy afternoon scheduling you may be seen in a timelier manner by accepting morning appointments.
            </p>
            <p style="text-align: justify; margin-bottom: 12px;">
              Our office strives to provide high-quality orthodontic treatment to patients to create a beautiful smile. To accomplish this, however, it will require a mutual commitment by both the dentist and the patient. Without the patient's cooperation, the completion of orthodontic treatment cannot be achieved in a timely manner or may never be achieved. Therefore, if the patient misses three (3) or more months in a twelve (12) month period, we will then assume our patient/orthodontic relationship has been terminated and that you will seek all future dental treatment at another dental office.
            </p>
            <div style="margin-top: 8px;">
              <strong>Initial:</strong> <span style="border-bottom: 2px solid #333; display: inline-block; min-width: 80px; padding-left: 5px;">${formData.initial5 || ''}</span>
            </div>
          </div>

          <!-- Section 6: Records and Discontinuation -->
          <div style="margin-bottom: 20px; font-size: 11px; line-height: 1.6; color: #495057;">
            <p style="text-align: justify; margin-bottom: 8px;">
              Should it be necessary to duplicate your records for transferable purposes due to re-location, we will forward a copy of your records within 10 working days to your new orthodontist for a fee of $100.
            </p>
            <p style="text-align: justify; margin-bottom: 8px;">
              If for any reason you discontinue treatment before completion, you must pay your previous delinquent monthly payment and debanding fee of $600. Additionally any discount previously given on the account will be retroactively voided, you will be responsible for the full amount of charges.
            </p>
            <p style="text-align: justify; margin-bottom: 12px;">
              Administration fee of $50 applies to all patients who are contracted for treatment but do not begin their treatment regardless of reason.
            </p>
            <div style="margin-top: 8px; margin-bottom: 20px;">
              <strong>Initial:</strong> <span style="border-bottom: 2px solid #333; display: inline-block; min-width: 80px; padding-left: 5px;">${formData.initial6 || ''}</span>
            </div>
            <p style="text-align: justify; margin-bottom: 12px;">
              It is understood and agreed that orthodontic treatment is elective in nature and all orthodontic appliances may be removed at any time without refund due to the following reasons: Non-payment of this account's financial obligations, excessive breakage (&gt;5) of orthodontic appliances, non-compliance, or poor oral hygiene.
            </p>
            <div style="margin-top: 8px;">
              <strong>Initial:</strong> <span style="border-bottom: 2px solid #333; display: inline-block; min-width: 80px; padding-left: 5px;">${formData.initial7 || ''}</span>
            </div>
          </div>
          
          <!-- Responsible Party Agreement -->
          <div style="margin-top: 20px; padding: 12px; border: 2px solid #666; background: #fafafa; page-break-inside: avoid;">
            
            <div style="margin-bottom: 12px;">
              <div style="margin-bottom: 4px; font-weight: 700; color: #666; font-size: 8px; text-transform: uppercase; letter-spacing: 0.3px;">Signature of Responsible Party:</div>
              ${termsSignatureData ? `
                <div style="border: 1.5px solid #666; padding: 6px; min-height: 80px; background: white;">
                  <img src="${termsSignatureData}" style="width: 100%; height: 70px; object-fit: contain;" />
                </div>
              ` : `
                <div style="border: 1.5px dashed #999; padding-top: 80px; background: white;"></div>
              `}
            </div>
            
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 12px;">
              <div>
                <div style="margin-bottom: 4px; font-weight: 700; color: #666; font-size: 8px; text-transform: uppercase; letter-spacing: 0.3px;">Name of Responsible Party:</div>
                <div style="border-bottom: 1.5px solid #666; padding: 6px 4px; min-height: 16px; background: white; font-weight: 500; color: #1a1a1a; font-size: 9px;">
                  ${formData.responsiblePartyName || ''}
                </div>
              </div>
              <div>
                <div style="margin-bottom: 4px; font-weight: 700; color: #666; font-size: 8px; text-transform: uppercase; letter-spacing: 0.3px;">Date:</div>
                <div style="border-bottom: 1.5px solid #666; padding: 6px 4px; min-height: 16px; background: white; font-weight: 500; color: #1a1a1a; font-size: 9px;">
                  ${formData.responsiblePartySignatureDate || ''}
                </div>
              </div>
            </div>
          </div>
        </div>
        
        <!-- Footer -->
        <div style="margin-top: 20px; padding: 8px 0; text-align: center; border-top: 2px solid #333; background: #f5f5f5;">
          <div style="font-size: 8px; font-weight: 700; color: #1a1a1a; margin-bottom: 3px; letter-spacing: 0.3px; text-transform: uppercase;">Smileland Dental</div>
          <div style="margin-top: 5px; padding-top: 5px; border-top: 1px solid #ccc;">
            <div style="font-size: 6px; color: #999; font-style: italic;">
              Contract Generated: ${new Date().toLocaleDateString('en-US', {
                year: 'numeric',
                month: 'long',
                day: 'numeric',
                hour: '2-digit',
                minute: '2-digit'
              })}
            </div>
          </div>
        </div>
      </body>
      </html>
    `;

    // Puppeteer로 PDF 생성
    const browser = await puppeteer.launch({
      headless: true,
      args: ['--no-sandbox', '--disable-setuid-sandbox']
    });
    
    const page = await browser.newPage();
    await page.setContent(html);
    
    // 페이지 로딩 대기
    await new Promise(resolve => setTimeout(resolve, 1000));
    
    const pdf = await page.pdf({
      format: 'A4',
      printBackground: true,
      margin: {
        top: '12mm',
        right: '12mm',
        bottom: '12mm',
        left: '12mm'
      },
      preferCSSPageSize: false,
      displayHeaderFooter: false
    });
    
    await browser.close();
    
    // 파일명 생성
    const patientName = formData.patientName?.replace(/\s+/g, '_') || 'Patient';
    const filename = `Orthodontic_Contract_${patientName}_${contractDate}.pdf`;
    
    return new NextResponse(Buffer.from(pdf), {
      headers: {
        'Content-Type': 'application/pdf',
        'Content-Disposition': `attachment; filename="${filename}"`
      }
    });

  } catch (error) {
    console.error('Orthodontic contract PDF generation error:', error);
    return NextResponse.json({ 
      success: false, 
      error: 'Failed to generate orthodontic contract PDF' 
    });
  }
}

