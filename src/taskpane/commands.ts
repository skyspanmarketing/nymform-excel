// Function file required by the manifest's ribbon button. Nymform has no ribbon commands
// beyond opening the task pane, so this only waits for Office.js.
import { ready } from "../office/adapter";

void ready().catch(() => undefined);
