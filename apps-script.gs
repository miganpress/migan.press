/**
 * 미간프레스 웹사이트 데이터 시트에 '사이트 업데이트' 메뉴를 추가합니다.
 * 누르면 GitHub에서 사이트 만들기·배포(.github/workflows/deploy.yml)를 바로 시작합니다.
 * (자동 갱신은 없습니다. 시트를 고친 뒤 이 메뉴를 눌러야 사이트에 반영됩니다.)
 *
 * 설치 방법
 * 1. GitHub에서 토큰을 만듭니다: 프로필 > Settings > Developer settings > Personal access tokens >
 *    Fine-grained tokens > Generate new token.
 *    Repository access는 이 저장소 하나만, Permissions > Repository permissions > Actions를 'Read and write'로 합니다.
 * 2. 시트에서 [확장 프로그램] > [Apps Script]를 엽니다. 이 파일 내용을 전부 붙여 넣고 저장합니다.
 * 3. 아래 REPO를 '계정이름/저장소이름'으로 고칩니다.
 * 4. 왼쪽 톱니바퀴(프로젝트 설정) > 스크립트 속성 > 속성 추가: 이름 GITHUB_TOKEN, 값은 1번의 토큰.
 *    (토큰을 코드에 직접 적지 마세요. 코드를 공유하면 토큰도 같이 노출됩니다.)
 * 5. 시트를 새로고침하면 위쪽 메뉴에 '미간 웹사이트'가 생깁니다.
 */
const REPO = '계정이름/저장소이름';
const WORKFLOW = 'deploy.yml';
const BRANCH = 'main';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('미간 웹사이트')
    .addItem('사이트 업데이트', 'updateSite')
    .addToUi();
}

function updateSite() {
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) {
    SpreadsheetApp.getUi().alert('스크립트 속성에 GITHUB_TOKEN이 없습니다. 파일 위쪽 설치 방법 4번을 확인하세요.');
    return;
  }
  const res = UrlFetchApp.fetch(
    'https://api.github.com/repos/' + REPO + '/actions/workflows/' + WORKFLOW + '/dispatches',
    {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28' },
      payload: JSON.stringify({ ref: BRANCH }),
      muteHttpExceptions: true,
    }
  );
  const ok = res.getResponseCode() === 204;
  SpreadsheetApp.getUi().alert(
    ok
      ? '업데이트를 시작했습니다. 1~2분 뒤 miganpress.com에 반영됩니다.'
      : '업데이트 요청이 실패했습니다 (' + res.getResponseCode() + '). REPO 이름과 토큰 권한을 확인하세요.'
  );
}
