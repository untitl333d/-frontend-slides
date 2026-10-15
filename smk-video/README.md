# SMK — промо-ролик (9:16, ~32 с)

Монтаж на [HyperFrames](https://github.com/heygen-com/hyperframes) в стиле кинетической типографики.

## Окружение

- Скиллы: `npx hyperframes skills update` (+ `faceless-explainer talking-head-recut motion-graphics embedded-captions`).
- CDN jsdelivr в облачном окружении закрыт, поэтому GSAP ставится из npm (`npm i gsap@3.14.2`)
  и подключается локально из `vendor/gsap.min.js`.
- Рендер: `HYPERFRAMES_SKIP_SKILLS=1 npx hyperframes render -o out.mp4`.

## Медиа

`assets/` в git не попадает (см. `.gitignore`): brand, trades, charts, telegram, sfx (Pixabay License), ref.
