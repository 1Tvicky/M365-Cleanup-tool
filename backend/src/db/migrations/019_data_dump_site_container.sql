-- Adds 'site' as a data_dump_containers.kind — a newly-created SharePoint site (spec §8/§9) is its
-- own container, distinct from 'site_library' (a document library *within* a site, existing or new).
ALTER TABLE data_dump_containers DROP CONSTRAINT data_dump_containers_kind_check;
ALTER TABLE data_dump_containers ADD CONSTRAINT data_dump_containers_kind_check
  CHECK (kind IN ('root_folder', 'folder', 'team', 'channel', 'site', 'site_library', 'mailbox_folder'));
