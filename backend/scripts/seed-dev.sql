-- Local dev only: one operator you can log in as immediately.
-- Email:    demo@cloudfuze.com
-- Password: Demo@12345
-- Hash generated with: node -e "require('bcryptjs').hash('Demo@12345', 12).then(console.log)"
INSERT INTO operators (email, display_name, password_hash, status)
VALUES (
  'demo@cloudfuze.com',
  'Demo Operator',
  '$2a$12$u7RYq2YxT4/.fDJdq/dlWuasef6JmDBK7Pba3IcM9GeNxbqJEOaYy',
  'active'
)
ON CONFLICT (email) DO UPDATE SET password_hash = EXCLUDED.password_hash, status = 'active';
