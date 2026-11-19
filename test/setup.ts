// Replace every network entry point with a function that throws. If any library
// code under test tries to reach the network, the test fails.
const deny = (name: string) => () => {
  throw new Error(`Network access attempted in test: ${name}`);
};

const g = globalThis as Record<string, unknown>;
g['fetch'] = deny('fetch');
g['XMLHttpRequest'] = function XMLHttpRequest() {
  deny('XMLHttpRequest')();
};
g['WebSocket'] = function WebSocket() {
  deny('WebSocket')();
};
g['EventSource'] = function EventSource() {
  deny('EventSource')();
};
if (typeof navigator !== 'undefined') {
  Object.defineProperty(navigator, 'sendBeacon', { value: deny('sendBeacon'), configurable: true });
}
