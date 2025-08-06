const slides = document.getElementById("slides");
let currentIndex = 0;

setInterval(() => {
	currentIndex = (currentIndex + 1) % 4;
	slides.style.transform = `translateX(-${currentIndex * 100}%)`;
}, 3000);

const hamburger = document.getElementById('hamburger');
const navWrapper = document.getElementById('nav-wrapper');
hamburger.addEventListener('click', () => {
	navWrapper.classList.toggle('open');
});

const scrollBtn = document.getElementById("scrollTopBtn");
scrollBtn.addEventListener("click", () => {
	window.scrollTo({ top: 0, behavior: "smooth" });
});

document.addEventListener("DOMContentLoaded", () => {
    const tabs = document.querySelectorAll('.tab');
    const contents = document.querySelectorAll('.tab-content');
    const moreLink = document.querySelector('.more');

    // 링크 매핑
    const links = {
        notice: "https://gmyouth.or.kr/nareum/selectBbsNttList.do?bbsNo=35&key=1247",
        recruit: "https://gmyouth.or.kr/nareum/selectBbsNttList.do?bbsNo=36&key=1248",
        community: "https://gmyouth.or.kr/nareum/selectBbsNttList.do?bbsNo=37&key=1249",
        data: "https://gmyouth.or.kr/nareum/selectBbsNttList.do?bbsNo=38&key=1250"
    };

    tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            tabs.forEach(t => t.classList.remove('active'));
            tab.classList.add('active');

            const target = tab.dataset.tab;
            contents.forEach(content => {
                if (content.dataset.content === target) {
                    content.style.display = 'block';
                } else {
                    content.style.display = 'none';
                }
            });

            // "더보기" 링크 업데이트
            moreLink.href = links[target];
        });
    });

    // 초기 설정
    const defaultTab = document.querySelector('.tab.active').dataset.tab;
    moreLink.href = links[defaultTab];
});

const date = new Date();
const today_year = date.getFullYear();
const today_month = date.getMonth() + 1;
const today_day = date.getDate();

// DOMContentLoaded 이벤트로 HTML 로딩 후 실행
document.addEventListener("DOMContentLoaded", function() {
  document.getElementById('activity_year').value = today_year;
  document.getElementById('activity_month').value = String(today_month).padStart(2, '0');
  document.getElementById('activity_day').value = String(today_day).padStart(2, '0');
});