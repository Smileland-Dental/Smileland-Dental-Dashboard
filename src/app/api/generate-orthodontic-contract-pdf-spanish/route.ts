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
      <html lang="es">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Contrato_de_Ortodoncia_${formData.patientName?.replace(/\s+/g, '_')}_${contractDate}</title>
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
          <h1>CONTRATO DE ORTODONCIA</h1>
          <p><strong>Smileland Dental</strong></p>
        </div>
        
        <div class="contract-info">
          <div class="info-row">
            <div class="info-item">
              <div class="info-label">Nombre d. Paciente</div>
              <div class="info-value">${formData.patientName || ''}</div>
            </div>
            <div class="info-item">
              <div class="info-label">Fecha d. Naci.</div>
              <div class="info-value">${formData.dob || ''}</div>
            </div>
          </div>
          
          <div class="info-row">
            <div class="info-item">
              <div class="info-label">Persona Responsable</div>
              <div class="info-value">${formData.responsibleParty || ''}</div>
            </div>
            <div class="info-item">
              <div class="info-label">Relacion al Paciente</div>
              <div class="info-value">${formData.relationship || ''}</div>
            </div>
          </div>
          
          <div class="info-row">
            <div class="info-item">
              <div class="info-label"># d. Seguro Social</div>
              <div class="info-value">${formData.ssn || ''}</div>
            </div>
            <div class="info-item">
              <div class="info-label"># Lic. d. Manejar</div>
              <div class="info-value">${formData.driversLicense || ''}</div>
            </div>
          </div>
        </div>
        
        <div class="treatment-section">
          <div class="treatment-label">TIPO DE TRATAMIENTO</div>
          <div class="treatment-options">
            <div class="treatment-option">
              <span class="checkbox ${formData.typeOfTreatment === 'Limited' ? 'checked' : ''}"></span>
              <span class="${formData.typeOfTreatment === 'Limited' ? 'selected' : ''}">Limitado</span>
            </div>
            <div class="treatment-option">
              <span class="checkbox ${formData.typeOfTreatment === 'Phase I' ? 'checked' : ''}"></span>
              <span class="${formData.typeOfTreatment === 'Phase I' ? 'selected' : ''}">Fase I</span>
            </div>
            <div class="treatment-option">
              <span class="checkbox ${formData.typeOfTreatment === 'Phase II' ? 'checked' : ''}"></span>
              <span class="${formData.typeOfTreatment === 'Phase II' ? 'selected' : ''}">Fase II</span>
            </div>
          </div>
        </div>
        
        ${formData.servicesRequired && formData.servicesRequired.length > 0 ? `
        <div class="services-section">
          <div class="section-header">Servicios Requeridos</div>
          <ul class="services-list">
            ${formData.servicesRequired.map((service: string) => {
              const [name, price] = service.split('|');
              return `<li>
                <span class="service-name">✓ ${name}</span>
                <span class="service-price">${price}</span>
              </li>`;
            }).join('')}
          </ul>
        </div>
        ` : ''}
        
        ${formData.additionalAppliances && formData.additionalAppliances.length > 0 ? `
        <div class="services-section">
          <div class="section-header">Aparatos (Si Necesario)</div>
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
            <h3 style="margin: 0 0 8px 0; text-align: center; color: #1a1a1a; font-size: 9px; font-weight: 700; letter-spacing: 0.4px; padding-bottom: 5px; border-bottom: 1.5px solid #333; text-transform: uppercase;">Primera Opicón</h3>
            
            <table style="width: 100%; border-collapse: collapse; font-size: 8px;">
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; width: 55%; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Total de Servicios:</td>
                <td style="padding: 3px 4px; color: #333; font-weight: 600; background: white; border: 1px solid #d0d0d0; border-left: none;">${formData.firstOption?.treatment ? '$' + formData.firstOption.treatment : ''}</td>
              </tr>
              <tr style="height: 2px;"></tr>
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Aparat:</td>
                <td style="padding: 3px 4px; color: #333; font-weight: 600; background: white; border: 1px solid #d0d0d0; border-left: none;">${formData.firstOption?.appliance ? '$' + formData.firstOption.appliance : ''}</td>
              </tr>
              <tr style="height: 2px;"></tr>
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Deposit:</td>
                <td style="padding: 3px 4px; color: #333; font-weight: 600; background: white; border: 1px solid #d0d0d0; border-left: none;">${formData.firstOption?.deposit ? '$' + formData.firstOption.deposit : ''}</td>
              </tr>
              <tr style="height: 2px;"></tr>
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Total Parcial:</td>
                <td style="padding: 3px 4px; color: #333; font-weight: 600; background: white; border: 1px solid #d0d0d0; border-left: none;">${formData.firstOption?.subtotal ? '$' + formData.firstOption.subtotal : ''}</td>
              </tr>
              <tr style="height: 2px;"></tr>
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Estimado de Seguro:</td>
                <td style="padding: 3px 4px; color: #333; font-weight: 600; background: white; border: 1px solid #d0d0d0; border-left: none;">${formData.firstOption?.estimatedInsurance ? '$' + formData.firstOption.estimatedInsurance : ''}</td>
              </tr>
              <tr style="height: 2px;"></tr>
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Balance Neto:</td>
                <td style="padding: 3px 4px; color: #333; font-weight: 600; background: white; border: 1px solid #d0d0d0; border-left: none;">${formData.firstOption?.netBalance ? '$' + formData.firstOption.netBalance : ''}</td>
              </tr>
              <tr style="height: 2px;"></tr>
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Periodo de Tratamiento Estimado:</td>
                <td style="padding: 3px 4px; color: #333; font-weight: 600; background: white; border: 1px solid #d0d0d0; border-left: none;">${formData.firstOption?.estimatedTreatmentPeriod ? '$' + formData.firstOption.estimatedTreatmentPeriod + ' meses' : ''}</td>
              </tr>
              <tr style="height: 2px;"></tr>
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Pago Mansual:</td>
                <td style="padding: 3px 4px; color: #333; font-weight: 600; background: white; border: 1px solid #d0d0d0; border-left: none;">${formData.firstOption?.monthlyPayment ? '$' + formData.firstOption.monthlyPayment : ''}</td>
              </tr>
            </table>
          </div>

          <!-- Second Option -->
          <div style="border: 1px solid #666; padding: 8px; background: #fafafa;">
            <h3 style="margin: 0 0 8px 0; text-align: center; color: #1a1a1a; font-size: 9px; font-weight: 700; letter-spacing: 0.4px; padding-bottom: 5px; border-bottom: 1.5px solid #333; text-transform: uppercase;">Segunda Opicón</h3>
            
            <table style="width: 100%; border-collapse: collapse; font-size: 8px;">
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; width: 55%; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Total de Servicios:</td>
                <td style="padding: 3px 4px; color: #333; font-weight: 600; background: white; border: 1px solid #d0d0d0; border-left: none;">${formData.secondOption?.treatment ? '$' + formData.secondOption.treatment : ''}</td>
              </tr>
              <tr style="height: 2px;"></tr>
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Aparat:</td>
                <td style="padding: 3px 4px; color: #333; font-weight: 600; background: white; border: 1px solid #d0d0d0; border-left: none;">${formData.secondOption?.appliance ? '$' + formData.secondOption.appliance : ''}</td>
              </tr>
              <tr style="height: 2px;"></tr>
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Deposit:</td>
                <td style="padding: 3px 4px; color: #333; font-weight: 600; background: white; border: 1px solid #d0d0d0; border-left: none;">${formData.secondOption?.deposit ? '$' + formData.secondOption.deposit : ''}</td>
              </tr>
              <tr style="height: 2px;"></tr>
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Total Parcial:</td>
                <td style="padding: 3px 4px; color: #333; font-weight: 600; background: white; border: 1px solid #d0d0d0; border-left: none;">${formData.secondOption?.subtotal ? '$' + formData.secondOption.subtotal : ''}</td>
              </tr>
              <tr style="height: 2px;"></tr>
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Estimado de Seguro:</td>
                <td style="padding: 3px 4px; color: #333; font-weight: 600; background: white; border: 1px solid #d0d0d0; border-left: none;">${formData.secondOption?.estimatedInsurance ? '$' + formData.secondOption.estimatedInsurance : ''}</td>
              </tr>
              <tr style="height: 2px;"></tr>
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Balance Neto:</td>
                <td style="padding: 3px 4px; color: #333; font-weight: 600; background: white; border: 1px solid #d0d0d0; border-left: none;">${formData.secondOption?.netBalance ? '$' + formData.secondOption.netBalance : ''}</td>
              </tr>
              <tr style="height: 2px;"></tr>
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Periodo de Tratamiento Estimado:</td>
                <td style="padding: 3px 4px; color: #333; font-weight: 600; background: white; border: 1px solid #d0d0d0; border-left: none;">${formData.secondOption?.estimatedTreatmentPeriod ? '$' + formData.secondOption.estimatedTreatmentPeriod + ' meses' : ''}</td>
              </tr>
              <tr style="height: 2px;"></tr>
              <tr>
                <td style="padding: 3px 4px; font-weight: 600; color: #1a1a1a; background: #f0f0f0; border: 1px solid #d0d0d0;">Pago Mansual:</td>
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
              <div style="margin-bottom: 3px; font-weight: 700; color: #666; font-size: 7px; text-transform: uppercase; letter-spacing: 0.2px;">Esitmado Presenta do Por:</div>
              <div style="border-bottom: 1px solid #666; padding: 5px 4px; min-height: 14px; background: white; font-weight: 500; color: #1a1a1a; font-size: 8px;">
                ${formData.quotePresentedBy || ''}
              </div>
            </div>
            <div>
              <div style="margin-bottom: 3px; font-weight: 700; color: #666; font-size: 7px; text-transform: uppercase; letter-spacing: 0.2px;">Fecha:</div>
              <div style="border-bottom: 1px solid #666; padding: 5px 4px; min-height: 14px; background: white; font-weight: 500; color: #1a1a1a; font-size: 8px;">
                ${formData.quotePresentedDate || ''}
              </div>
            </div>
          </div>
          
          <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 8px;">
            <div>
              <div style="margin-bottom: 3px; font-weight: 700; color: #666; font-size: 7px; text-transform: uppercase; letter-spacing: 0.2px;">Firma de Persona Responsable:</div>
              ${signatureData ? `
                <div style="border: 1px solid #666; padding: 5px; min-height: 60px; background: white;">
                  <img src="${signatureData}" style="width: 100%; height: 52px; object-fit: contain;" />
                </div>
              ` : `
                <div style="border: 1px dashed #999; padding-top: 60px; background: white;"></div>
              `}
            </div>
            <div>
              <div style="margin-bottom: 3px; font-weight: 700; color: #666; font-size: 7px; text-transform: uppercase; letter-spacing: 0.2px;">Fecha:</div>
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
              El saldo pendiente de $ <span style="border-bottom: 1px solid #333; display: inline-block; min-width: 60px; text-align: center; padding: 0 5px;">${formData.unpaidBalance || '_______'}</span> será pagado en <span style="border-bottom: 1px solid #333; display: inline-block; min-width: 40px; text-align: center; padding: 0 5px;">${formData.paymentMonths || '_____'}</span> meses en la cantidad de $ <span style="border-bottom: 1px solid #333; display: inline-block; min-width: 60px; text-align: center; padding: 0 5px;">${formData.monthlyPaymentAmount || '_______'}</span>, cada uno, por el día 1o o 15 de cada mes, comenzando él <span style="border-bottom: 1px solid #333; display: inline-block; min-width: 80px; text-align: center; padding: 0 5px;">${formData.paymentBeginDate || '__________'}</span> y continuando hasta que el saldo anterior sea pagado en su totalidad. Un cargo de $20.00 será aplicado a su cuenta si su pago no se recibe en 5 días de la fecha de vencimiento Smileland acepta dinero en efectivo, tarjetas de crédito y Care Credit como forma de pago. Cheques personales podrán ser aceptados solo por correo.
            </p>
            <p style="text-align: justify; margin-bottom: 12px;">
              El saldo pendiente se ha dividido en pagos mensuales solamente como ayuda del saldo total y no esta correlacionada con sus citas. El costo total debe ser pagado por completo antes de retirar los aparatos de ortodoncia y la colocación de los retenedores. Aparatos de retención se pondrán a la conclusión del tratamiento ortodontico. Si la cuenta llega a estar atrasada, el paciente no será visto para ningún tratamiento solo para emergencia de tratamiento paliativo sólo para aliviar dolor con $50 cargo. Si su cuenta está en atraso más de 3 meses, el caso será terminado y será referido a la agencia de la colección y el honorario adicional será de $600 cargado a la cuenta.
            </p>
            <div style="margin-top: 8px;">
              <strong>Iniciales:</strong> <span style="border-bottom: 2px solid #333; display: inline-block; min-width: 80px; padding-left: 5px;">${formData.initial1 || ''}</span>
            </div>
          </div>

          <!-- Section 2: Insurance -->
          <div style="margin-bottom: 20px; font-size: 11px; line-height: 1.6; color: #495057;">
            <p style="text-align: justify; margin-bottom: 8px;">
              Si usted tiene cobertura de ortodoncia por favor entienda que la cobertura de seguro es un contrato entre usted y su compañía de seguros. Nosotros le ayudaremos en la facturación de su compañía de seguros. Si su compañía de seguros no paga su parte, será su responsabilidad el seguimiento con ellos directamente. Usted es personalmente responsable del costo total de tratamiento recibido. Si el seguro no paga la cuenta del tratamiento estimado, usted será responsable del saldo.
            </p>
            <p style="text-align: justify; margin-bottom: 8px;">
              Aunque después de Facturar el seguro dental, el seguro puede enviar pagos al suscriptor (a usted o al asegurado principal) por lo tanto es su responsabilidad de traer el pago a la oficina para eliminar su saldo pendiente. Si el pago se envía bajo el nombre de los suscriptores, es posible que tenga cobrar el cheque primero y luego traer la cantidad a nuestra oficina. Toda la cantidad pagada al suscriptor y no pagada hacia la cuenta del paciente será considerado delincuente y por lo tanto pueden ser referidos a una agencia de cobranzas.
            </p>
            <p style="text-align: justify; margin-bottom: 12px;">
              Si el paciente es menor de 18 años, los padres, padre o tutor serán legalmente responsables de cualquier saldo pendiente.
            </p>
            <div style="margin-top: 8px;">
              <strong>Iniciales:</strong> <span style="border-bottom: 2px solid #333; display: inline-block; min-width: 80px; padding-left: 5px;">${formData.initial2 || ''}</span>
            </div>
          </div>

          <!-- Section 3: General Dentistry -->
          <div style="margin-bottom: 20px; font-size: 11px; line-height: 1.6; color: #495057;">
            <p style="text-align: justify; margin-bottom: 12px;">
              Los Cargos De tratamiento de ortodoncia no incluyen ningún costo de odontología general, tales como rellenos, exponer y adherir, extracciones, coronas o limpiezas. Los chequeos dentales regulares y todo el trabajo dental necesario deben ser completados y documentados por su dentista antes de la adhesión de ortodoncia. Por favor pida que su dentista consulte con nosotros si requiere cualquier tratamiento dental extenso. La buena higiene oral es esencial para el éxito del tratamiento de ortodoncia. Si el paciente requiere un periodo de tratamiento adicional debido al mal higiene bucal, no usare elástico, y faltas de citas, un cargo de $100 por cada mes será agregado a su cuenta.
            </p>
            <div style="margin-top: 8px;">
              <strong>Iniciales:</strong> <span style="border-bottom: 2px solid #333; display: inline-block; min-width: 80px; padding-left: 5px;">${formData.initial3 || ''}</span>
            </div>
          </div>

          <!-- Section 4: Broken Appliances -->
          <div style="margin-bottom: 20px; font-size: 11px; line-height: 1.6; color: #495057;">
            <p style="text-align: justify; margin-bottom: 12px;">
              El reemplazo de aplicaciones pérdidas o rotas resultara en un costo adicional más el costo del laboratorio. El costo de los soportes de metal es $35 y los soportes claros son $45 por cada uno.
            </p>
            <div style="margin-top: 8px;">
              <strong>Iniciales:</strong> <span style="border-bottom: 2px solid #333; display: inline-block; min-width: 80px; padding-left: 5px;">${formData.initial4 || ''}</span>
            </div>
          </div>

          <!-- Section 5: Missed Appointments -->
          <div style="margin-bottom: 20px; font-size: 11px; line-height: 1.6; color: #495057;">
            <p style="text-align: justify; margin-bottom: 8px;">
              Trataremos de alojarle al hacer sus citas. <strong>Cualquier mes perdido o cancelación sin un aviso de 24 horas será sujeta a un cargo de $20 por cita; por lo tanto es importante que nos notifique inmediatamente si tiene que cambiar su cita.</strong> Será responsable de contactar a nuestra oficina y reprogramar cualquier cita perdida. Si llega más de 15 minutos tarde a su cita, es probable que tuviera que reprogramar su cita. Debido a muchas citas por la tarde, usted puede ser visto en un tiempo más razonable si acepta citas en la mañana.
            </p>
            <p style="text-align: justify; margin-bottom: 12px;">
              Nuestra oficina se esfuerza para proveer tratamiento de ortodoncia de alta calidad a pacientes para crear una sonrisa hermosa. Para llevar a cabo esto, requerirá un compromiso mutuo tanto del dentista como del paciente. Sin la cooperación de los pacientes, la finalización del tratamiento de ortodoncia no puede ser conseguida en una manera oportuna o nunca podría ser lograda. Por lo tanto, si el paciente pierde tres (3) o más meses en un período de doce (12) meses, supondremos entonces que nuestra relación de paciente y oficina de ortodoncia ha sido terminada y que buscará todo el futuro tratamiento dental en otro consultorio dental.
            </p>
            <div style="margin-top: 8px;">
              <strong>Iniciales:</strong> <span style="border-bottom: 2px solid #333; display: inline-block; min-width: 80px; padding-left: 5px;">${formData.initial5 || ''}</span>
            </div>
          </div>

          <!-- Section 6: Records and Discontinuation -->
          <div style="margin-bottom: 20px; font-size: 11px; line-height: 1.6; color: #495057;">
            <p style="text-align: justify; margin-bottom: 8px;">
              Si es necesario duplicar sus expedientes para los propósitos de la transferencia debido a la relocalización, remitiremos una copia de sus expedientes dentro de 10 días laborables a su nuevo ortodontista con un honorario de $100.00.
            </p>
            <p style="text-align: justify; margin-bottom: 8px;">
              Si por alguna razón usted interrumpe el tratamiento antes de que todos los pagos se hayan hecho, usted debe pagar su antiguo balance y una cuota de $600 para quitar aparatos. El cargo administrativo de $50.00 se aplica a todos los pacientes que firmen su contrato para el tratamiento y por algún motivo no se presenten para comenzar su tratamiento.
            </p>
            <p style="text-align: justify; margin-bottom: 12px;">
              Cargo de administración de $50 se aplicara a todos los pacientes que son contratados para recibir tratamiento, pero no empiezan su tratamiento, por cualquier razón.
            </p>
            <div style="margin-top: 8px; margin-bottom: 20px;">
              <strong>Iniciales:</strong> <span style="border-bottom: 2px solid #333; display: inline-block; min-width: 80px; padding-left: 5px;">${formData.initial6 || ''}</span>
            </div>
            <p style="text-align: justify; margin-bottom: 12px;">
              Se entiende y acuerda que el tratamiento de ortodoncia es de naturaleza electiva y todos los aparatos ortodonticos pueden ser removidos en cualquier momento sin reembolso debido a las siguientes razones: La falta de pago de las obligaciones financieras de esta cuenta, la rotura excesiva (&gt;5) de los aparatos ortodonticos, la falta de cumplimiento, o pobres medidas de higiene oral que serian perjudiciales para la función opcional y/o estética.
            </p>
            <div style="margin-top: 8px;">
              <strong>Iniciales:</strong> <span style="border-bottom: 2px solid #333; display: inline-block; min-width: 80px; padding-left: 5px;">${formData.initial7 || ''}</span>
            </div>
          </div>
          
          <!-- Responsible Party Agreement -->
          <div style="margin-top: 20px; padding: 12px; border: 2px solid #666; background: #fafafa; page-break-inside: avoid;">
            
            <div style="margin-bottom: 12px;">
              <div style="margin-bottom: 4px; font-weight: 700; color: #666; font-size: 8px; text-transform: uppercase; letter-spacing: 0.3px;">Firma de Persona Responsable:</div>
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
                <div style="margin-bottom: 4px; font-weight: 700; color: #666; font-size: 8px; text-transform: uppercase; letter-spacing: 0.3px;">Nombre de Persona Responsable:</div>
                <div style="border-bottom: 1.5px solid #666; padding: 6px 4px; min-height: 16px; background: white; font-weight: 500; color: #1a1a1a; font-size: 9px;">
                  ${formData.responsiblePartyName || ''}
                </div>
              </div>
              <div>
                <div style="margin-bottom: 4px; font-weight: 700; color: #666; font-size: 8px; text-transform: uppercase; letter-spacing: 0.3px;">Fecha:</div>
                <div style="border-bottom: 1.5px solid #666; padding: 6px 4px; min-height: 16px; background: white; font-weight: 500; color: #1a1a1a; font-size: 9px;">
                  ${formData.responsiblePartySignatureDate || ''}
                </div>
              </div>
            </div>
          </div>
        </div>
        
        <!-- Footer -->
        <div style="margin-top: 20px; padding: 8px 0; text-align: center; border-top: 2px solid #333; background: #f5f5f5;">
          <div style="font-size: 8px; font-weight: 700; color: #1a1a1a; margin-bottom: 3px; letter-spacing: 0.3px; text-transform: uppercase;">SMILELAND DENTAL</div>
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
    const filename = `Contrato_de_Ortodoncia_${patientName}_${contractDate}.pdf`;
    
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

