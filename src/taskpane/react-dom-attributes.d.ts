// Chrome's writing suggestions can be turned off per field with the writingsuggestions attribute,
// which React's types don't know yet. React passes it to the element as written.
import "react";

declare module "react" {
  // The type parameter has to match React's declaration.
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  interface HTMLAttributes<T> {
    writingsuggestions?: "true" | "false";
  }
}
