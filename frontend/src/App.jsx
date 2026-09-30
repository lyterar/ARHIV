import { useState } from "react";
import { registerStudent } from "./api.js";
import TicketCard from "./TicketCard.jsx";

export default function App() {
  const [lastName, setLastName] = useState("");
  const [firstName, setFirstName] = useState("");
  const [ticket, setTicket] = useState(null);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  async function submit(e) {
    e.preventDefault();
    if (!lastName.trim() || !firstName.trim()) {
      setError("Значение не может быть пустым.");
      return;
    }
    setError("");
    setLoading(true);
    try {
      setTicket(await registerStudent(lastName, firstName));
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  }

  function reset() {
    setTicket(null);
    setLastName("");
    setFirstName("");
  }

  return (
    <main className="page">
      <div className="stage">
        {ticket ? (
          <TicketCard key={ticket.ticket + ticket.text} ticket={ticket} onReset={reset} />
        ) : (
          <form className="panel tilt" onSubmit={submit} noValidate>
            <h1>Получить билет</h1>
            <p className="muted">Введите фамилию и имя — билет выпадет случайно.</p>
            <label>
              Фамилия
              <input value={lastName} onChange={(e) => setLastName(e.target.value)} autoFocus />
            </label>
            <label>
              Имя
              <input value={firstName} onChange={(e) => setFirstName(e.target.value)} />
            </label>
            {error && <p className="error">{error}</p>}
            <button type="submit" disabled={loading}>
              {loading ? "Тянем билет…" : "Взять билет"}
            </button>
          </form>
        )}
      </div>
    </main>
  );
}
