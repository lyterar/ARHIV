import { useEffect, useRef, useState } from "react";

// A 3D card that flips over on mount and follows the pointer.
export default function TicketCard({ ticket, onReset }) {
  const cardRef = useRef(null);
  const [flipped, setFlipped] = useState(false);
  const [copied, setCopied] = useState(false);

  useEffect(() => {
    const t = setTimeout(() => setFlipped(true), 150);
    return () => clearTimeout(t);
  }, []);

  function onMove(e) {
    const el = cardRef.current;
    const r = el.getBoundingClientRect();
    const x = (e.clientX - r.left) / r.width - 0.5;
    const y = (e.clientY - r.top) / r.height - 0.5;
    el.style.setProperty("--ry", `${x * 14}deg`);
    el.style.setProperty("--rx", `${-y * 14}deg`);
  }

  function onLeave() {
    cardRef.current.style.setProperty("--ry", "0deg");
    cardRef.current.style.setProperty("--rx", "0deg");
  }

  async function copy() {
    try {
      await navigator.clipboard.writeText(ticket.text);
      setCopied(true);
      setTimeout(() => setCopied(false), 1500);
    } catch {
      /* clipboard unavailable */
    }
  }

  return (
    <div className="card-wrap" onPointerMove={onMove} onPointerLeave={onLeave}>
      <div ref={cardRef} className="card-tilt">
        <div className={`card ${flipped ? "flipped" : ""}`}>
          <div className="face front">
            <span className="q">?</span>
          </div>
          <div className="face back">
            <span className="badge">Билет</span>
            <div className="number">{ticket.ticket}</div>
            <pre className="ticket-text">{ticket.text}</pre>
          </div>
        </div>
      </div>
      <div className={`actions ${flipped ? "show" : ""}`}>
        <button onClick={copy}>{copied ? "Скопировано" : "Копировать текст"}</button>
        <button className="ghost" onClick={onReset}>
          Следующий студент
        </button>
      </div>
    </div>
  );
}
