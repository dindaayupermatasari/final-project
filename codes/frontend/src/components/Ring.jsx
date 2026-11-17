import { warnaSkor } from "../utils.js";

// Cincin skor: hijau >= 80, kuning 60-79, merah di bawahnya. `color` menimpa warna cincin.
export default function Ring({ score, className = "", color }) {
  const c = color || warnaSkor(score);
  return (
    <div className={`ring ${className}`.trim()} style={{ "--v": score, "--c": c }}>
      <span>{score}%</span>
    </div>
  );
}
