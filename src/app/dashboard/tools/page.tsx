'use client';

import { useState, useEffect } from 'react';

export default function DoctorPage() {
  const [model, setModel] = useState('34');
  const [customPercent, setCustomPercent] = useState(34);
  const [daysPerYear, setDaysPerYear] = useState(260);
  const [slowDays, setSlowDays] = useState(60);
  const [busyDays, setBusyDays] = useState(200);
  const [retentionBonus, setRetentionBonus] = useState(100000);
  const [bonusYears, setBonusYears] = useState(3);
  const [teamIncentive, setTeamIncentive] = useState(65);

  // Experienced DDS
  const [expUseSeasons, setExpUseSeasons] = useState(false);
  const [expDaily, setExpDaily] = useState(4300);
  const [expSlow, setExpSlow] = useState(3500);
  const [expBusy, setExpBusy] = useState(4800);

  // Medium DDS
  const [medUseSeasons, setMedUseSeasons] = useState(false);
  const [medDaily, setMedDaily] = useState(3500);
  const [medSlow, setMedSlow] = useState(3000);
  const [medBusy, setMedBusy] = useState(3900);

  // New DDS
  const [newUseSeasons, setNewUseSeasons] = useState(false);
  const [newDaily, setNewDaily] = useState(2500);
  const [newSlow, setNewSlow] = useState(2200);
  const [newBusy, setNewBusy] = useState(2800);

  const [results, setResults] = useState({
    experienced: { prod: 0, bonus: 0, team: 0, total: 0 },
    medium: { prod: 0, bonus: 0, team: 0, total: 0 },
    new: { prod: 0, bonus: 0, team: 0, total: 0 },
  });

  function getSelectedPercent() {
    if (model === 'custom') {
      return customPercent;
    }
    return parseFloat(model);
  }

  function formatMoney(num: number) {
    if (!isFinite(num)) return '-';
    return num.toLocaleString('en-US', {
      minimumFractionDigits: 0,
      maximumFractionDigits: 0,
    });
  }

  function calculateForDDS(
    useSeasons: boolean,
    single: number,
    slow: number,
    busy: number
  ) {
    const percent = getSelectedPercent() / 100;
    const annualBonus = retentionBonus / bonusYears;
    const teamIncentiveTotal = teamIncentive * daysPerYear;

    let annualProductionPay = 0;

    if (useSeasons) {
      const slowPay = slow * slowDays * percent;
      const busyPay = busy * busyDays * percent;
      annualProductionPay = slowPay + busyPay;
    } else {
      annualProductionPay = single * daysPerYear * percent;
    }

    const totalComp = annualProductionPay + annualBonus + teamIncentiveTotal;

    return {
      prod: annualProductionPay,
      bonus: annualBonus,
      team: teamIncentiveTotal,
      total: totalComp,
    };
  }

  useEffect(() => {
    const expResult = calculateForDDS(expUseSeasons, expDaily, expSlow, expBusy);
    const medResult = calculateForDDS(medUseSeasons, medDaily, medSlow, medBusy);
    const newResult = calculateForDDS(newUseSeasons, newDaily, newSlow, newBusy);

    setResults({
      experienced: expResult,
      medium: medResult,
      new: newResult,
    });
  }, [
    model,
    customPercent,
    daysPerYear,
    slowDays,
    busyDays,
    retentionBonus,
    bonusYears,
    teamIncentive,
    expUseSeasons,
    expDaily,
    expSlow,
    expBusy,
    medUseSeasons,
    medDaily,
    medSlow,
    medBusy,
    newUseSeasons,
    newDaily,
    newSlow,
    newBusy,
  ]);

  const currentPercent = getSelectedPercent();

  return (
    <>
      <style jsx global>{`
        * {
          box-sizing: border-box;
          font-family: system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
        }

        .doctor-container {
          max-width: 1100px;
          margin: 0 auto;
          padding: 24px;
          background: #f5f5f7;
          min-height: 100vh;
          color: #222;
        }

        .doctor-container h1 {
          margin-top: 0;
          font-size: 1.8rem;
        }

        .doctor-container h2 {
          font-size: 1.3rem;
          margin-bottom: 0.4rem;
        }

        .doctor-card {
          background: #ffffff;
          border-radius: 16px;
          padding: 20px;
          margin-bottom: 18px;
          box-shadow: 0 8px 20px rgba(0, 0, 0, 0.06);
        }

        .doctor-grid {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
          gap: 16px;
        }

        .doctor-card label {
          display: block;
          font-size: 0.9rem;
          margin-bottom: 4px;
        }

        .doctor-card input[type='number'],
        .doctor-card input[type='text'] {
          width: 100%;
          padding: 8px 10px;
          border-radius: 8px;
          border: 1px solid #d0d0d5;
          font-size: 0.95rem;
        }

        .doctor-card input[type='number']:focus,
        .doctor-card input[type='text']:focus {
          outline: none;
          border-color: #0070f3;
          box-shadow: 0 0 0 1px #0070f3;
        }

        .radio-group {
          display: flex;
          flex-wrap: wrap;
          gap: 12px;
          margin-bottom: 8px;
        }

        .radio-pill {
          display: flex;
          align-items: center;
          gap: 6px;
          padding: 6px 10px;
          border-radius: 999px;
          border: 1px solid #d0d0d5;
          background: #fafafa;
          cursor: pointer;
          font-size: 0.9rem;
        }

        .radio-pill input {
          margin: 0;
        }

        .section-title {
          font-weight: 600;
          margin-bottom: 6px;
        }

        .small-text {
          font-size: 0.8rem;
          color: #666;
        }

        .toggle-row {
          display: flex;
          align-items: center;
          gap: 8px;
          margin-top: 8px;
          font-size: 0.85rem;
        }

        .btn-primary {
          border: none;
          border-radius: 999px;
          padding: 10px 18px;
          font-size: 0.95rem;
          cursor: pointer;
          background: #0070f3;
          color: white;
        }

        .btn-primary:hover {
          background: #005ad0;
        }

        .results-table {
          width: 100%;
          border-collapse: collapse;
          margin-top: 8px;
          font-size: 0.9rem;
        }

        .results-table th,
        .results-table td {
          border: 1px solid #e0e0e5;
          padding: 8px 10px;
          text-align: right;
        }

        .results-table th:first-child,
        .results-table td:first-child {
          text-align: left;
        }

        .results-table th {
          background: #fafafa;
          font-weight: 600;
        }

        .tag {
          display: inline-block;
          padding: 2px 8px;
          border-radius: 999px;
          font-size: 0.75rem;
          background: #e8f0ff;
          color: #174ea6;
          margin-left: 6px;
        }

        @media (max-width: 600px) {
          .doctor-container {
            padding: 16px;
          }
        }
      `}</style>

      <div className="doctor-container">
        <h1>Calculator</h1>
        <p className="small-text">
          Use this calculator to estimate annual compensation for restorative DDS under different compensation percentages.
          You can plug in average daily production and customize assumptions (days/year, retention bonus, team incentive, etc.).
        </p>

        {/* Global settings */}
        <div className="doctor-card">
          <h2>1. Global Settings</h2>

          <div className="section-title">Compensation Model</div>
          <div className="radio-group">
            <label className="radio-pill">
              <input
                type="radio"
                name="model"
                value="34"
                checked={model === '34'}
                onChange={(e) => setModel(e.target.value)}
              />
              <span>34% (Final Model)</span>
            </label>
            <label className="radio-pill">
              <input
                type="radio"
                name="model"
                value="40"
                checked={model === '40'}
                onChange={(e) => setModel(e.target.value)}
              />
              <span>40% (Trial Scenario)</span>
            </label>
            <label className="radio-pill">
              <input
                type="radio"
                name="model"
                value="custom"
                checked={model === 'custom'}
                onChange={(e) => setModel(e.target.value)}
              />
              <span>Custom %</span>
            </label>
            <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
              <label htmlFor="customPercent" className="small-text">
                Custom %
              </label>
              <input
                type="number"
                id="customPercent"
                value={customPercent}
                onChange={(e) => setCustomPercent(parseFloat(e.target.value) || 0)}
                min="0"
                max="100"
                step="0.1"
                style={{ width: '80px' }}
              />
            </div>
          </div>

          <div className="doctor-grid" style={{ marginTop: '12px' }}>
            <div>
              <label htmlFor="daysPerYear">Total working days per year</label>
              <input
                type="number"
                id="daysPerYear"
                value={daysPerYear}
                onChange={(e) => setDaysPerYear(parseFloat(e.target.value) || 0)}
                min="1"
              />
              <div className="small-text">Includes PTO + paid holidays.</div>
            </div>
            <div>
              <label htmlFor="slowDays">Slow season days</label>
              <input
                type="number"
                id="slowDays"
                value={slowDays}
                onChange={(e) => setSlowDays(parseFloat(e.target.value) || 0)}
                min="0"
              />
            </div>
            <div>
              <label htmlFor="busyDays">Busy season days</label>
              <input
                type="number"
                id="busyDays"
                value={busyDays}
                onChange={(e) => setBusyDays(parseFloat(e.target.value) || 0)}
                min="0"
              />
            </div>
            <div>
              <label htmlFor="retentionBonus">Retention bonus total ($)</label>
              <input
                type="number"
                id="retentionBonus"
                value={retentionBonus}
                onChange={(e) => setRetentionBonus(parseFloat(e.target.value) || 0)}
                min="0"
              />
            </div>
            <div>
              <label htmlFor="bonusYears">Retention bonus period (years)</label>
              <input
                type="number"
                id="bonusYears"
                value={bonusYears}
                onChange={(e) => setBonusYears(parseFloat(e.target.value) || 1)}
                min="1"
              />
            </div>
            <div>
              <label htmlFor="teamIncentive">Team incentive per day ($)</label>
              <input
                type="number"
                id="teamIncentive"
                value={teamIncentive}
                onChange={(e) => setTeamIncentive(parseFloat(e.target.value) || 0)}
                min="0"
              />
            </div>
          </div>
        </div>

        {/* DDS inputs */}
        <div className="doctor-card">
          <h2>2. DDS Inputs</h2>
          <p className="small-text">
            For each DDS type, enter either a single average daily production number or separate slow/busy values.
            The calculator will apply the selected % model (e.g., 34% or 40%) to estimate their pay.
          </p>

          <div className="doctor-grid">
            {/* Experienced */}
            <div className="dds-card" data-row="experienced">
              <div className="section-title">
                Experienced DDS <span className="tag">Example</span>
              </div>

              <div className="toggle-row">
                <input
                  type="checkbox"
                  id="expUseSeasons"
                  checked={expUseSeasons}
                  onChange={(e) => setExpUseSeasons(e.target.checked)}
                />
                <label htmlFor="expUseSeasons">Use separate slow/busy daily production</label>
              </div>

              {!expUseSeasons && (
                <div className="exp-single">
                  <label htmlFor="expDaily">Average daily production ($)</label>
                  <input
                    type="number"
                    id="expDaily"
                    value={expDaily}
                    onChange={(e) => setExpDaily(parseFloat(e.target.value) || 0)}
                    min="0"
                  />
                </div>
              )}

              {expUseSeasons && (
                <div className="doctor-grid exp-seasonal" style={{ marginTop: '8px' }}>
                  <div>
                    <label htmlFor="expSlow">Slow season daily production ($)</label>
                    <input
                      type="number"
                      id="expSlow"
                      value={expSlow}
                      onChange={(e) => setExpSlow(parseFloat(e.target.value) || 0)}
                      min="0"
                    />
                  </div>
                  <div>
                    <label htmlFor="expBusy">Busy season daily production ($)</label>
                    <input
                      type="number"
                      id="expBusy"
                      value={expBusy}
                      onChange={(e) => setExpBusy(parseFloat(e.target.value) || 0)}
                      min="0"
                    />
                  </div>
                </div>
              )}
            </div>

            {/* Medium */}
            <div className="dds-card" data-row="medium">
              <div className="section-title">Medium DDS</div>

              <div className="toggle-row">
                <input
                  type="checkbox"
                  id="medUseSeasons"
                  checked={medUseSeasons}
                  onChange={(e) => setMedUseSeasons(e.target.checked)}
                />
                <label htmlFor="medUseSeasons">Use separate slow/busy daily production</label>
              </div>

              {!medUseSeasons && (
                <div className="med-single">
                  <label htmlFor="medDaily">Average daily production ($)</label>
                  <input
                    type="number"
                    id="medDaily"
                    value={medDaily}
                    onChange={(e) => setMedDaily(parseFloat(e.target.value) || 0)}
                    min="0"
                  />
                </div>
              )}

              {medUseSeasons && (
                <div className="doctor-grid med-seasonal" style={{ marginTop: '8px' }}>
                  <div>
                    <label htmlFor="medSlow">Slow season daily production ($)</label>
                    <input
                      type="number"
                      id="medSlow"
                      value={medSlow}
                      onChange={(e) => setMedSlow(parseFloat(e.target.value) || 0)}
                      min="0"
                    />
                  </div>
                  <div>
                    <label htmlFor="medBusy">Busy season daily production ($)</label>
                    <input
                      type="number"
                      id="medBusy"
                      value={medBusy}
                      onChange={(e) => setMedBusy(parseFloat(e.target.value) || 0)}
                      min="0"
                    />
                  </div>
                </div>
              )}
            </div>

            {/* New */}
            <div className="dds-card" data-row="new">
              <div className="section-title">New DDS</div>

              <div className="toggle-row">
                <input
                  type="checkbox"
                  id="newUseSeasons"
                  checked={newUseSeasons}
                  onChange={(e) => setNewUseSeasons(e.target.checked)}
                />
                <label htmlFor="newUseSeasons">Use separate slow/busy daily production</label>
              </div>

              {!newUseSeasons && (
                <div className="new-single">
                  <label htmlFor="newDaily">Average daily production ($)</label>
                  <input
                    type="number"
                    id="newDaily"
                    value={newDaily}
                    onChange={(e) => setNewDaily(parseFloat(e.target.value) || 0)}
                    min="0"
                  />
                </div>
              )}

              {newUseSeasons && (
                <div className="doctor-grid new-seasonal" style={{ marginTop: '8px' }}>
                  <div>
                    <label htmlFor="newSlow">Slow season daily production ($)</label>
                    <input
                      type="number"
                      id="newSlow"
                      value={newSlow}
                      onChange={(e) => setNewSlow(parseFloat(e.target.value) || 0)}
                      min="0"
                    />
                  </div>
                  <div>
                    <label htmlFor="newBusy">Busy season daily production ($)</label>
                    <input
                      type="number"
                      id="newBusy"
                      value={newBusy}
                      onChange={(e) => setNewBusy(parseFloat(e.target.value) || 0)}
                      min="0"
                    />
                  </div>
                </div>
              )}
            </div>
          </div>

          <div style={{ marginTop: '16px' }}>
            <span className="small-text" style={{ marginLeft: '10px' }}>
              Current model: {currentPercent.toFixed(1)}% of restorative production
            </span>
          </div>
        </div>

        {/* Results */}
        <div className="doctor-card">
          <h2>3. Results – Estimated Annual Compensation</h2>
          <table className="results-table">
            <thead>
              <tr>
                <th>DDS Type</th>
                <th>Annual Production Pay ($)</th>
                <th>Retention Bonus (Annualized)</th>
                <th>Team Incentive</th>
                <th>Total Annual Compensation ($)</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td>Experienced DDS</td>
                <td>{formatMoney(results.experienced.prod)}</td>
                <td>{formatMoney(results.experienced.bonus)}</td>
                <td>{formatMoney(results.experienced.team)}</td>
                <td>{formatMoney(results.experienced.total)}</td>
              </tr>
              <tr>
                <td>Medium DDS</td>
                <td>{formatMoney(results.medium.prod)}</td>
                <td>{formatMoney(results.medium.bonus)}</td>
                <td>{formatMoney(results.medium.team)}</td>
                <td>{formatMoney(results.medium.total)}</td>
              </tr>
              <tr>
                <td>New DDS</td>
                <td>{formatMoney(results.new.prod)}</td>
                <td>{formatMoney(results.new.bonus)}</td>
                <td>{formatMoney(results.new.team)}</td>
                <td>{formatMoney(results.new.total)}</td>
              </tr>
            </tbody>
          </table>
          <p className="small-text" style={{ marginTop: '8px' }}>
            Note: This tool assumes restorative DDS are paid based on the selected percentage of their daily production.
            Exam DDS can be modeled separately using their own daily production and percentage.
          </p>
        </div>
      </div>
    </>
  );
}
