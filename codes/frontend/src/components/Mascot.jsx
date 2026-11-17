import idle from "../assets/img/mascot-idle.png";
import happy from "../assets/img/mascot-happy.png";
import wink from "../assets/img/mascot-wink.png";
import surprised from "../assets/img/mascot-surprised.png";
import sleepy from "../assets/img/mascot-sleepy.png";
import confused from "../assets/img/mascot-confused.png";
import search from "../assets/img/mascot-search.png";

const IMG = { idle, happy, wink, surprised, sleepy, confused, search };

// Pose: idle | happy | wink | surprised | sleepy | confused | search. Badge: teks kecil di pojok (mis. "!", "?", "✓", "↑").
export default function Mascot({ pose = "idle", size = 96, badge, tone, style, className = "", ...rest }) {
  return (
    <span
      className={`mi ${pose} ${className}`.trim()}
      style={{
        "--s": `${size}px`,
        width: `${size}px`,
        height: `${size}px`,
        aspectRatio: "1 / 1",
        backgroundImage: `url(${IMG[pose] || IMG.idle})`,
        backgroundSize: "contain",
        backgroundPosition: "center",
        backgroundRepeat: "no-repeat",
        ...style,
      }}
      {...rest}
    >
      {badge && <b className={`bd ${tone === "blue" ? "b" : ""}`.trim()}>{badge}</b>}
    </span>
  );
}
