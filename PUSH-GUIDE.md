# GitHubへ初回pushする手順

このZIPはGit履歴を含みません。展開後に内容を確認し、新しいPrivateリポジトリとして登録してください。

## 0. 事前確認

1. 会社・部門のルール上、ソースコードをGitHubへ保存してよいか確認する
2. GitHub上のリポジトリは **Private** にする
3. ZIPを任意の作業フォルダーへ展開する
4. `ManualBuilder` フォルダー内に実際の業務画像、個人情報、パスワード、トークンがないことを確認する
5. Windowsへ [Git for Windows](https://git-scm.com/download/win) をインストールする

## 1. GitHubで空のリポジトリを作る

GitHubの `New repository` から次の内容で作成します。

- Repository name: `ManualBuilder`
- Visibility: `Private`
- Add a README file: オフ
- Add .gitignore: None
- Choose a license: None

README、`.gitignore`、ライセンスはZIP内に用意済みです。GitHub側では空のリポジトリを作成します。

## 2. PowerShellから初回コミットする

PowerShellを開き、展開したフォルダーへ移動します。パスは実際の保存先に合わせてください。

```powershell
cd "$env:USERPROFILE\Downloads\ManualBuilder"

git --version
git init -b main

git config user.name "あなたの名前"
git config user.email "GitHubに登録したメールアドレス"

git status
git add .
git status
git commit -m "chore: initialize ManualBuilder repository"
```

2回目の `git status` で、次が登録対象に含まれていないことを確認してください。

- 実際の業務スクリーンショット
- `.xlsx`、`.xlsm`、`.docx`、`.pdf`
- `out`、`.tmp`、`data`、`projects`、`screenshots`、`exports`
- パスワード、アクセストークン、秘密鍵、`.env`

間違ったファイルが表示された場合は、その時点で `git commit` を実行せず、ファイルを移動するか `.gitignore` を修正します。

## 3. GitHubへ接続してpushする

GitHubで作った空リポジトリのHTTPS URLへ置き換えて実行します。

```powershell
git remote add origin https://github.com/YOUR_ACCOUNT/ManualBuilder.git
git remote -v
git push -u origin main
```

認証画面が出た場合は、Git Credential Managerによるブラウザー認証を使う方法が簡単です。コマンドラインでパスワードを直接求められる環境では、GitHubのアカウントパスワードではなくPersonal Access Tokenが必要です。トークンをスクリプトやリポジトリへ保存しないでください。

## 4. Phase 0の確定タグを付ける

初回pushの後、検証済みの基準点としてタグを付けます。

```powershell
git tag -a phase0-v2.16 -m "Phase 0 validation completed"
git push origin phase0-v2.16
```

GitHubの画面で、`main` ブランチと `phase0-v2.16` タグの両方が表示されれば完了です。

## GitHub CLIを使う場合

GitHub CLIが入っている場合は、GitHub上のリポジトリ作成とpushをまとめられます。`git commit` まで済ませた後に実行します。

```powershell
gh auth login
gh repo create ManualBuilder --private --source=. --remote=origin --push
git tag -a phase0-v2.16 -m "Phase 0 validation completed"
git push origin phase0-v2.16
```

## 以後の基本運用

Phase 1は作業ブランチで進めます。

```powershell
git switch -c phase1/mvp-foundation

# ファイルを編集した後
git status
git add src web docs
git commit -m "feat: add project foundation"
git push -u origin phase1/mvp-foundation
```

GitHub上でPull Requestを作り、内容を確認してから `main` に取り込みます。

## push後の確認

- リポジトリがPrivateになっている
- `tests/phase0` のスクリプトが表示される
- `out` や生成Officeファイルが表示されていない
- 実際の業務画像や個人情報が表示されていない
- `phase0-v2.16` タグが作成されている

参考: [既存のローカルコードをGitHubへ追加する公式手順](https://docs.github.com/en/migrations/importing-source-code/using-the-command-line-to-import-source-code/adding-locally-hosted-code-to-github)、[GitHubのコマンドライン認証](https://docs.github.com/en/authentication/keeping-your-account-and-data-secure/about-authentication-to-github)
