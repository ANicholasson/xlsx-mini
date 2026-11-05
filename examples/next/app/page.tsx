import { ExportButton } from "./export-button";
import { orders } from "./orders";

export default function Home() {
  return (
    <main style={{ fontFamily: "system-ui, sans-serif", padding: "2rem", maxWidth: "40rem" }}>
      <h1>xlsx-mini in Next.js</h1>
      <p>Two ways to hand a workbook to the browser, both from the same {orders.length} orders.</p>
      <ul>
        <li>
          <a href="/api/report">Download from the route handler</a>: the server returns the file as
          a response.
        </li>
        <li>
          <ExportButton />: a server action returns the bytes and the client saves them.
        </li>
      </ul>
    </main>
  );
}
