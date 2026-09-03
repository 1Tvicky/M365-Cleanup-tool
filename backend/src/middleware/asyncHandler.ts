import type { NextFunction, Request, RequestHandler, Response } from "express";

/**
 * Express 4 does not catch a rejected promise thrown from an async route handler — an unhandled
 * `throw new ApiError(...)` inside one just hangs the request instead of reaching errorHandler.
 * Wrap every async handler with this so thrown/rejected errors reach app.ts's error middleware.
 */
export function asyncHandler(
  fn: (req: Request, res: Response, next: NextFunction) => Promise<unknown>
): RequestHandler {
  return (req, res, next) => {
    fn(req, res, next).catch(next);
  };
}
