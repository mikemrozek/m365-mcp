// unpdf's type declarations import types from `@napi-rs/canvas`, which is an
// OPTIONAL peer dependency used only for rendering PDF pages to images. We
// extract text and never render, so that package is deliberately not installed
// (it is a ~40MB platform-specific native binding, which would complicate the
// container build for no benefit).
//
// This stub satisfies the type resolver for that one module. Preferred over
// `skipLibCheck: true`, which would suppress genuine type errors across every
// dependency in the tree rather than just this one deliberate omission.
// Declared as types (not a bare module) because unpdf uses these names in type
// positions; a bare `declare module` makes them namespaces and fails to compile.
declare module '@napi-rs/canvas' {
  export type Canvas = unknown;
  export type SKRSContext2D = unknown;
}
