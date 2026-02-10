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

  const lineRes = await dbQuery(
    `INSERT INTO production_lines(name, secret_key, created_by_user_id)
     VALUES ($1, $2, $3)
     ON CONFLICT (name)
     DO UPDATE SET updated_at = NOW()
     RETURNING id, name`,
    ['Production Line 1', 'SEED-KEY-001', manager.id]
  );
  const line = lineRes.rows[0];

  await dbQuery(
    `INSERT INTO supervisor_line_assignments(supervisor_user_id, line_id, assigned_by_user_id)
     VALUES ($1, $2, $3)
     ON CONFLICT (supervisor_user_id, line_id) DO NOTHING`,
    [supervisor.id, line.id, manager.id]
  );

  console.log('Seed complete');
  console.log(`Manager login: manager / manager123`);
  console.log(`Supervisor login: supervisor / supervisor123`);
}

run()
  .catch((error) => {
    console.error(error);
    process.exitCode = 1;
  })
  .finally(async () => {
    await pool.end();
  });
