-- Backs the "Name" field in the Manage Clouds expand-panel summary (the connecting admin's
-- display name, distinct from admin_upn which is their sign-in identifier).
ALTER TABLE connections ADD COLUMN admin_display_name TEXT;
