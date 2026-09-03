import pg from "pg";
import { config } from "../config/index.js";

export const pool = new pg.Pool({ connectionString: config.database.url });

export async function query<T extends pg.QueryResultRow = pg.QueryResultRow>(
  text: string,
  params: unknown[] = []
): Promise<pg.QueryResult<T>> {
  return pool.query<T>(text, params);
}
