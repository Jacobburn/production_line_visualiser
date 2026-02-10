import { dbQuery, pool } from '../src/db.js';
import { hashPassword } from '../src/auth.js';

async function upsertUser({ name, username, password, role }) {
  const passwordHash = await hashPassword(password);
  const result = await dbQuery(
    `INSERT INTO users(name, username, password_hash, role, is_active)
     VALUES ($1, $2, $3, $4, TRUE)
     ON CONFLICT (username)
     DO UPDATE SET name = EXCLUDED.name, password_hash = EXCLUDED.password_hash, role = EXCLUDED.role, is_active = TRUE
     RETURNING id, name, username, role`,
    [name, username.toLowerCase(), passwordHash, role]
  );
  return result.rows[0];
}

async function run() {
  const manager = await upsertUser({
    name: 'Manager',
    username: 'manager',
    password: 'manager123',
    role: 'manager'
  });

  const supervisor = await upsertUser({
    name: 'Supervisor',
    username: 'supervisor',
    password: 'supervisor123',
    role: 'supervisor'
  });

  const baseLineRes = await dbQuery(
    `INSERT INTO production_lines(name, secret_key, created_by_user_id)
     VALUES ($1, $2, $3)
     ON CONFLICT (name)
     DO UPDATE SET updated_at = NOW()
     RETURNING id, name`,
    ['Production Line 1', 'SEED-KEY-001', manager.id]
  );
  const baseLine = baseLineRes.rows[0];

  const demoLineRes = await dbQuery(
    `INSERT INTO production_lines(name, secret_key, created_by_user_id)
     VALUES ($1, $2, $3)
     ON CONFLICT (name)
     DO UPDATE SET updated_at = NOW()
     RETURNING id, name`,
    ['Demo #1', 'DEMO-KEY-001', manager.id]
  );
  const demoLine = demoLineRes.rows[0];

  for (const lineId of [baseLine.id, demoLine.id]) {
    await dbQuery(
      `INSERT INTO supervisor_line_assignments(supervisor_user_id, line_id, assigned_by_user_id)
       VALUES ($1, $2, $3)
       ON CONFLICT (supervisor_user_id, line_id) DO NOTHING`,
      [supervisor.id, lineId, manager.id]
    );
  }

  const stageSeed = [
    { order: 1, name: '1. Tumbler', type: 'prep', dayCrew: 2, nightCrew: 2, max: 12, x: 77, y: 66, w: 14, h: 24 },
    { order: 2, name: '2. Transfer', type: 'transfer', dayCrew: 2, nightCrew: 2, max: 10, x: 55, y: 80, w: 12, h: 8 },
    { order: 3, name: '3. Kebab Box Pack', type: 'prep', dayCrew: 4, nightCrew: 4, max: 12, x: 24, y: 58, w: 20, h: 18 },
    { order: 4, name: '4. Kebab Box Cut', type: 'prep', dayCrew: 1, nightCrew: 1, max: 12, x: 11, y: 58, w: 10, h: 16 },
    { order: 5, name: '5. Kebab Box Unload', type: 'prep', dayCrew: 1, nightCrew: 1, max: 12, x: 2, y: 58, w: 9, h: 16 },
    { order: 6, name: '6. Transfer', type: 'transfer', dayCrew: 1, nightCrew: 1, max: 10, x: 2, y: 42, w: 9, h: 12 },
    { order: 7, name: '7. Kebab Split', type: 'main', dayCrew: 1, nightCrew: 1, max: 12, x: 2, y: 14, w: 10, h: 16 },
    { order: 8, name: '8. Marinate', type: 'main', dayCrew: 5, nightCrew: 8, max: 12, x: 12, y: 14, w: 18, h: 16 },
    { order: 9, name: '9. Wipe', type: 'main', dayCrew: 1, nightCrew: 1, max: 12, x: 31, y: 14, w: 9, h: 14 },
    { order: 10, name: '10. Proseal', type: 'main', dayCrew: 1, nightCrew: 1, max: 12, x: 40, y: 14, w: 12, h: 16 },
    { order: 11, name: '11. Metal Detector', type: 'main', dayCrew: 1, nightCrew: 1, max: 12, x: 55, y: 16, w: 8, h: 14 },
    { order: 12, name: '12. Bottom Labeller', type: 'main', dayCrew: 1, nightCrew: 1, max: 12, x: 66, y: 14, w: 11, h: 16 },
    { order: 13, name: '13. Top Labeller', type: 'main', dayCrew: 1, nightCrew: 1, max: 12, x: 78, y: 14, w: 10, h: 16 },
    { order: 14, name: '14. Pack', type: 'main', dayCrew: 2, nightCrew: 2, max: 12, x: 88, y: 15, w: 8, h: 16 }
  ];

  await dbQuery(`DELETE FROM line_stages WHERE line_id = $1`, [demoLine.id]);
  await dbQuery(`DELETE FROM line_layout_guides WHERE line_id = $1`, [demoLine.id]);

  const stageRows = [];
  for (const stage of stageSeed) {
    const inserted = await dbQuery(
      `INSERT INTO line_stages(
        line_id, stage_order, stage_name, stage_type, day_crew, night_crew, max_throughput_per_crew, x, y, w, h
      ) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11)
      RETURNING id, stage_order`,
      [
        demoLine.id,
        stage.order,
        stage.name,
        stage.type,
        stage.dayCrew,
        stage.nightCrew,
        stage.max,
        stage.x,
        stage.y,
        stage.w,
        stage.h
      ]
    );
    stageRows.push(inserted.rows[0]);
  }
  const stageIdsByOrder = new Map(stageRows.map((row) => [Number(row.stage_order), row.id]));

  await dbQuery(`DELETE FROM shift_logs WHERE line_id = $1`, [demoLine.id]);
  await dbQuery(`DELETE FROM run_logs WHERE line_id = $1`, [demoLine.id]);
  await dbQuery(`DELETE FROM downtime_logs WHERE line_id = $1`, [demoLine.id]);

  const start = new Date(2026, 0, 1);
  const days = 60;
  const dayProducts = ['Teriyaki', 'Honey Soy', 'Sweet Chilli'];
  const nightProducts = ['Peri Peri', 'Lemon Herb', 'Hot & Spicy'];
  const dayEquipmentOrders = [8, 10, 12, 13, 14, 7, 9];
  const nightEquipmentOrders = [3, 4, 5, 10, 11, 12, 14];

  for (let i = 0; i < days; i += 1) {
    const d = new Date(start);
    d.setDate(start.getDate() + i);
    const date = d.toISOString().slice(0, 10);
    const isWeekend = d.getDay() === 0 || d.getDay() === 6;

    const dayCrew = isWeekend ? 16 : 18;
    const nightCrew = isWeekend ? 14 : 16;

    await dbQuery(
      `INSERT INTO shift_logs(
         line_id, date, shift, crew_on_shift, start_time, break1_start, break2_start, break3_start, finish_time, submitted_by_user_id
       ) VALUES ($1,$2,'Day',$3,'06:00','09:00','12:00','14:00','14:00',$4)`,
      [demoLine.id, date, dayCrew, supervisor.id]
    );

    await dbQuery(
      `INSERT INTO shift_logs(
         line_id, date, shift, crew_on_shift, start_time, break1_start, break2_start, break3_start, finish_time, submitted_by_user_id
       ) VALUES ($1,$2,'Night',$3,'14:00','17:00','20:00','22:00','22:00',$4)`,
      [demoLine.id, date, nightCrew, supervisor.id]
    );

    const dayP1 = dayProducts[i % dayProducts.length];
    const dayP2 = dayProducts[(i + 1) % dayProducts.length];
    const nightP1 = nightProducts[i % nightProducts.length];
    const nightP2 = nightProducts[(i + 1) % nightProducts.length];

    await dbQuery(
      `INSERT INTO run_logs(
         line_id, date, shift, product, setup_start_time, production_start_time, finish_time, units_produced, submitted_by_user_id
       ) VALUES ($1,$2,'Day',$3,'05:40','06:10','10:35',$4,$5)`,
      [demoLine.id, date, dayP1, 2800 + (i % 7) * 90, supervisor.id]
    );

    await dbQuery(
      `INSERT INTO run_logs(
         line_id, date, shift, product, setup_start_time, production_start_time, finish_time, units_produced, submitted_by_user_id
       ) VALUES ($1,$2,'Day',$3,'10:40','10:55','15:25',$4,$5)`,
      [demoLine.id, date, dayP2, 2500 + (i % 5) * 110, supervisor.id]
    );

    await dbQuery(
      `INSERT INTO run_logs(
         line_id, date, shift, product, setup_start_time, production_start_time, finish_time, units_produced, submitted_by_user_id
       ) VALUES ($1,$2,'Night',$3,'13:55','14:15','18:55',$4,$5)`,
      [demoLine.id, date, nightP1, 2400 + (i % 6) * 95, supervisor.id]
    );

    await dbQuery(
      `INSERT INTO run_logs(
         line_id, date, shift, product, setup_start_time, production_start_time, finish_time, units_produced, submitted_by_user_id
       ) VALUES ($1,$2,'Night',$3,'21:10','21:30','00:25',$4,$5)`,
      [demoLine.id, date, nightP2, 2150 + (i % 8) * 85, supervisor.id]
    );

    await dbQuery(
      `INSERT INTO downtime_logs(
         line_id, date, shift, downtime_start, downtime_finish, equipment_stage_id, reason, submitted_by_user_id
       ) VALUES ($1,$2,'Day','08:10',$3,$4,'Planned clean-down',$5)`,
      [
        demoLine.id,
        date,
        `08:${String(22 + (i % 8)).padStart(2, '0')}`,
        stageIdsByOrder.get(dayEquipmentOrders[i % dayEquipmentOrders.length]) || null,
        supervisor.id
      ]
    );

    await dbQuery(
      `INSERT INTO downtime_logs(
         line_id, date, shift, downtime_start, downtime_finish, equipment_stage_id, reason, submitted_by_user_id
       ) VALUES ($1,$2,'Day','11:20',$3,$4,'Label alignment',$5)`,
      [
        demoLine.id,
        date,
        `11:${String(30 + (i % 10)).padStart(2, '0')}`,
        stageIdsByOrder.get(dayEquipmentOrders[(i + 2) % dayEquipmentOrders.length]) || null,
        supervisor.id
      ]
    );

    await dbQuery(
      `INSERT INTO downtime_logs(
         line_id, date, shift, downtime_start, downtime_finish, equipment_stage_id, reason, submitted_by_user_id
       ) VALUES ($1,$2,'Night','16:30',$3,$4,'Sensor reset',$5)`,
      [
        demoLine.id,
        date,
        `16:${String(42 + (i % 9)).padStart(2, '0')}`,
        stageIdsByOrder.get(nightEquipmentOrders[i % nightEquipmentOrders.length]) || null,
        supervisor.id
      ]
    );

    await dbQuery(
      `INSERT INTO downtime_logs(
         line_id, date, shift, downtime_start, downtime_finish, equipment_stage_id, reason, submitted_by_user_id
       ) VALUES ($1,$2,'Night','20:05',$3,$4,'Film change',$5)`,
      [
        demoLine.id,
        date,
        `20:${String(18 + (i % 11)).padStart(2, '0')}`,
        stageIdsByOrder.get(nightEquipmentOrders[(i + 3) % nightEquipmentOrders.length]) || null,
        supervisor.id
      ]
    );
  }

  console.log('Seed complete');
  console.log(`Manager login: manager / manager123`);
  console.log(`Supervisor login: supervisor / supervisor123`);
  console.log(`Demo line with sample data: Demo #1`);
}

run()
  .catch((error) => {
    console.error(error);
    process.exitCode = 1;
  })
  .finally(async () => {
    await pool.end();
  });
