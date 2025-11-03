# FAX-Automation-Tool

![Excel VBA](https://img.shields.io/badge/-Excel%20VBA-217346?logo=microsoft-excel&logoColor=white)
![Office Automation](https://img.shields.io/badge/-Office%20Automation-4CAF50)
![RPA](https://img.shields.io/badge/-RPA-FF9800)
![Portfolio](https://img.shields.io/badge/-Portfolio-black)

Excel上の依頼データからFAX送信用の原本を自動生成し、複数事業所への転送・印刷を一括化するVBAツール。

# 📠 提供表FAX送付状 自動作成ツール (Excel VBA)  
**Automated FAX Cover Sheet Generator for Care Service Providers**

---

## 🧭 概要 / Overview

このツールは、Excel VBA を使って **「サービスチェックシート」から各事業所ごとのFAX送付状を自動生成** する仕組みです。  
介護・医療系の業務で、複数宛先に同じ書類を送る際の手間を大幅に削減します。  

This Excel VBA tool automatically generates individual FAX cover sheets for each care office  
based on a master sheet ("サービスチェックシート"). It is designed to streamline FAX preparation in care or medical operations.

---

## ⚙️ 主な機能 / Key Features

✅ **ダブルクリックで自動生成**  
Just double-click on the sheet to start generation.

✅ **事業所ごとに自動シート作成**  
Each care office gets its own sheet cloned from a FAX template.

✅ **利用者名の重複除去・整列**  
Automatically removes duplicate client names and formats the list neatly.

✅ **FAX送信枚数を自動計算**  
Auto-calculates total pages to be sent (count × 2 + 1).

✅ **安全なシート名変換**  
Automatically removes invalid characters and trims names for Excel compliance.

---

## 🧩 シート構成 / Sheet Structure

| シート名 | 役割 | Description |
|:--|:--|:--|
| サービスチェックシート | 元データ（A列＝事業所名、B列＝利用者名） | Base data |
| FAX原本 | テンプレートシート | Template sheet |
| 自動生成された各シート | 各事業所ごとのFAX送付状 | Generated sheets |

---

## 🔍 動作の流れ / Process Flow

1. 「サービスチェックシート」のA列（事業所名）とB列（利用者名）を走査  
2. 「居宅介護支援事業所しらゆりケア」を除外  
3. 同一事業所名の利用者をグループ化・重複除去  
4. 「FAX原本」を複製し以下を出力：  
   - A9：事業所名  
   - A11：利用者名リスト（4名ごとに改行）  
   - E5：送信枚数（利用者数×2+1）  
5. メッセージ「シートの作成が完了しました。」を表示

---

## 🧠 コード構成 / VBA Logic Overview

**主要イベント：**
```vb
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
