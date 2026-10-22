# 财务RPA 加载宏构建工具（VBA-compiler）

将纯文本 `.bas` 源代码一键构建为 WPS 加载宏（`.xlam`），并可打包成分发安装包。
本仓库不是"编译器"，而是 **源代码 → 加载宏 → 安装包** 的构建/分发流水线 + 一组财务数据清洗宏。

> 完整使用与维护说明见 **[操作手册.md](操作手册.md)**。

## 目录结构

```
.
├── src/                 源代码（UTF-8，唯一真相，日常只改这里）
│   ├── ModuleCommon.bas 公共工具：分类映射 / 表头校验 / 伴生文件定位 / SafeCDbl·Nz
│   ├── Module2.bas      库存价值表清洗
│   ├── Module3.bas      库龄表清洗
│   ├── Module4.bas      库龄期初变化分析
│   ├── Module5.bas      采购发票清洗
│   ├── Module6.bas      固定资产隐藏列
│   ├── Module7.bas      工具栏创建（AutoSetupToolbar）+ 关于
│   └── ThisWorkbook.cls 启动事件（参考）
├── templates/           打包用模板（install.vbs / 使用说明）
├── build.vbs            构建：src → FinanceRPA.xlam（编码归一/硬失败校验/部署）
├── fix_isaddin.ps1      构建子步骤：修补 xlam 内 workbook.xml 的 isAddin
├── pack.vbs             打包：生成 FinanceRPA安装包/
└── 操作手册.md          详细操作手册
```

## 三步口诀

```
改代码 → 关 WPS → 双击 build.vbs
双击 pack.vbs → 生成安装包
复制安装包 → 双击 install.vbs
```

## 编码约定

- `src/*.bas`、`templates/*` 一律 **UTF-8（无 BOM）**，由 `.gitattributes` 锁定。
- `build.vbs` 只在临时目录 `_build_gbk/` 生成 GBK 副本用于导入，**绝不改写 `src/`**。
