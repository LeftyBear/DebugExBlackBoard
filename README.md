# VBA開発 正典

## 1. アーキテクチャ

- DDD / Clean Architecture を採用する
- 依存方向は外側 → 内側のみ
- Domain層は他層を参照しない
- Application層はDomain層のみ参照可能
- Infrastructure層はDomain/Applicationを利用可能
- Presentation層はApplicationを利用する

---

## 2. 依存関係

### 正しい依存方向

Pre_View
↓
App_Factory
↓
App_UseCase
↓
Pre_Presenter
↓
Pre_ViewModel

### 禁止事項

- Pre_Presenter → Pre_View
- App_UseCase → Pre_View
- Domain → Application
- Domain → Infrastructure
- Pre_ViewModel → Pre_View

---

## 3. Presenterの責務

### Presenterが行うこと

- UseCase結果をViewModelへ変換
- 画面表示用データの整形
- ViewModel生成

### Presenterが行わないこと

- View保持
- Control操作
- MsgBox表示
- UserForm操作

### 正典

Pre_PresenterはPre_View参照を保持しない。

---

## 4. Viewの責務

### Viewが保持するもの

- App_Factory

### Viewが保持しないもの

- Presenter
- Repository
- Entity

### 正典

Pre_View → App_Factory → App_UseCase

で処理を開始する。

---

## 5. UseCaseの責務

### UseCaseが行うこと

- アプリケーションルール実行
- Domain呼び出し
- Repository呼び出し
- Presenter呼び出し

### UseCaseが行わないこと

- Control操作
- UserForm操作
- MsgBox表示

---

## 6. ViewModel

### 目的

Pre_Presenterが生成する画面表示専用オブジェクト

### 特徴

- ロジックを持たない
- 画面描画情報のみ保持
- Presentation層専用

---

## 7. クラス命名規約

### Domain

```vb
Dom_Entity
Dom_ValueObject
Dom_Service
```

### Application

```vb
App_UseCase
App_Factory
```

### Presentation

```vb
Pre_View
Pre_Form
Pre_Control
Pre_Presenter
Pre_ViewModel
```

### Infrastructure

```vb
Inf_Repository
Inf_Dao
Inf_Gateway
```

---

## 8. 命名規約

### クラス

- PascalCase

例

```vb
App_CustomerSearchUseCase
Pre_CustomerListPresenter
Dom_Customer
```

### メソッド

- PascalCase

例

```vb
Execute
CreateViewModel
FindById
```

### 変数

- PascalCase

例

```vb
CustomerId
CustomerList
ViewModel
```

---

## 9. 引数規約

### 原則

すべて明示的に ByVal を付与する

```vb
Public Function Execute(ByVal CustomerId As String) As Pre_ViewModel
```

### 例外

オブジェクト参照を変更する場合のみ ByRef を使用する

---

## 10. Selector / Resolver 分離

### Selector

選択処理を担当

```vb
SelectCustomer
SelectSheet
```

### Resolver

解決処理を担当

```vb
ResolveCustomer
ResolveRepository
```

### 正典

SelectorとResolverの責務を混在させない。

---

## 11. Mapping / Dictionary

### 正典

Application層でDictionaryを利用しない。

### 理由

- 暗黙的なマッピングを防ぐ
- 可読性向上
- 保守性向上

### 推奨

明示的なMapperクラスを作成する。

```vb
App_CustomerMapper
```

---

## 12. クラス内部状態

### 必須パターン

```vb
Private Type Member

End Type

Private This As Member
```

### 参照方法

```vb
This.CustomerId
This.CustomerName
```

---

## 13. VBAコーディング規約

### 行継続

使用しない。

非推奨

```vb
Execute _
    Value
```

推奨

```vb
Execute Value
```

### Call

使用しない。

非推奨

```vb
Call Execute(Value)
```

推奨

```vb
Execute Value
```

### 可読性

- ネストは浅くする
- Exit FunctionよりGuard節を優先する
- 意味のある名前を使用する

---

## 14. Repository

### 役割

永続化抽象化

### UseCaseから見えるもの

```vb
ICustomerRepository
```

### Infrastructure実装

```vb
Inf_CustomerRepository
```

### 正典

UseCaseはInterfaceへ依存する。

---

## 15. Factory

### Viewが保持するもの

```vb
App_Factory
```

### Factoryの責務

- UseCase生成
- Presenter生成
- Repository注入

### 正典

依存関係構築はFactoryへ集約する。

---

## 16. UI制御

### UserForm

表示専用

### Control

Pre_Viewからのみ操作する。

### 禁止

```vb
App_UseCase → TextBox
Pre_Presenter → ListBox
Dom_Entity → UserForm
```

---

## 17. 最終正典

依存関係

Pre_View
↓
App_Factory
↓
App_UseCase
↓
Pre_Presenter
↓
Pre_ViewModel

Pre_PresenterはPre_Viewを保持しない。

Pre_ViewはApp_Factoryのみ保持する。

App_UseCaseはPre_Presenterを利用してPre_ViewModelを生成する。

Pre_ViewはPre_ViewModelを描画する。

Control操作はPre_Viewのみが行う。
