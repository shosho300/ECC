# CSSD Central Hub — نسخة Microsoft SharePoint

<div dir="rtl">

هذا المجلد يحوّل موقع **CSSD Central Hub** (اللي كان ملف HTML واحد على chatgpt.site) إلى **Web Part في SharePoint Online** مبني بـ SharePoint Framework (SPFx).

الواجهة والأقسام نفسها: الطلبات الخارجية، استلام الطلبات، البلاغات، الإنشارج، المجموعات المستعارة، إدارة القسم، مكتبة IFU، الحالات الطارئة، والريكول. الفرق إن البيانات صارت **مشتركة بين كل الموظفين ومحفوظة في SharePoint** بدل ما تكون محفوظة في متصفح كل شخص.

## وش اللي تغيّر عن النسخة القديمة

| النسخة القديمة (HTML) | نسخة SharePoint |
|---|---|
| البيانات في `localStorage` داخل متصفح كل جهاز | البيانات في **قوائم SharePoint** على الموقع، يشوفها كل من له صلاحية |
| المرفقات مخزنة داخل المتصفح (حد 2 MB) | المرفقات ترفع إلى مكتبة **CSSD Attachments** (حد 10 MB) |
| ما فيه فصل صلاحيات | الصلاحيات من SharePoint: تقدرين تقفلين قوائم الإدارة على مجموعة الإنشارج فقط |
| بيانات تجريبية جاهزة | يبدأ فاضي، بيانات حقيقية فقط |
| اسم المستخدم يُكتب يدويًا | يتعبى تلقائيًا من حساب Microsoft 365 (ويمكن تعديله) |

## القوائم اللي تنشأ على الموقع

| القائمة | المحتوى |
|---|---|
| CSSD Orders | طلبات العيادات والعمليات واستلامها |
| CSSD Issues | البلاغات ومتابعتها |
| CSSD Loaner Sets | المجموعات المستعارة |
| CSSD Shift Handovers | تسليم الشفت |
| CSSD Parameters | سجلات البراميترس |
| CSSD Staff Tasks | مهام الستاف (من صفحة الإدارة) |
| CSSD Device Status | جاهزية الأجهزة |
| CSSD IFU Library | مكتبة تعليمات الاستخدام |
| CSSD Emergency Cases | الحالات الطارئة |
| CSSD Recalls | الريكول وردود العمليات |
| CSSD Attachments (مكتبة مستندات) | كل الملفات المرفقة |

كل قائمة فيها أعمدة واضحة (القسم، التاريخ، الشفت، الحالة…) عشان تقدرين تفتحينها في SharePoint أو Excel أو Power BI مباشرة. والسجل الكامل محفوظ في عمود **CSSD record (JSON)**.

## طريقة التركيب (مرة وحدة)

تحتاجين حساب **SharePoint Administrator** أو صلاحية على **App Catalog**.

1. **ارفعي الحزمة إلى App Catalog**
   - افتحي SharePoint Admin Center ← **More features** ← **Apps** ← **App Catalog**.
   - ارفعي الملف [`package/cssd-central-hub.sppkg`](package/cssd-central-hub.sppkg).
   - لما تطلع رسالة التأكيد اختاري **Enable app** (يمكن تفعيله لكل المواقع مباشرة).
2. **أنشئي موقع للقسم** (إذا ما عندك): مثلًا Communication Site باسم `CSSD`.
3. **أضيفي التطبيق للموقع**: من الموقع ← ⚙️ ← **Add an app** ← اختاري **cssd-central-hub**.
4. **أنشئي صفحة**: ⚙️ ← **Add a page**، اختاري تخطيط **Full-width section**، وأضيفي Web Part اسمه **CSSD Central Hub**، وبعدها **Publish**.
5. **أول مرة تفتحين الصفحة** (بحساب مالك الموقع): بتطلع لك رسالة إن القوائم غير موجودة. اضغطي **Create SharePoint lists** وتنشأ كل القوائم والمكتبة تلقائيًا.
6. **الصلاحيات**:
   - موظفين العيادات والعمليات والتعقيم: صلاحية **Edit** (Members) على الموقع.
   - إذا تبين قوائم الإدارة (CSSD Staff Tasks و CSSD Device Status) أو قرار الريكول تكون للإنشارج بس: افتحي القائمة ← ⚙️ ← **List settings** ← **Permissions for this list** وأوقفي الوراثة، وعطي الصلاحية للمجموعة المطلوبة.

> نصيحة: خلي الصفحة هي الصفحة الرئيسية للموقع (Make homepage) وأضيفي الموقع كـ Tab في Microsoft Teams لقناة القسم. الـ Web Part يشتغل داخل Teams بعد.

## ملاحظات مهمة

- **البيانات القديمة**: النسخة القديمة كانت تحفظ في متصفح كل جهاز، ولا يمكن نقلها تلقائيًا. السجلات تبدأ من جديد في SharePoint.
- **تعارض التعديل**: إذا شخصين عدلوا نفس السجل بنفس الوقت، الثاني يطلع له تنبيه إن السجل تغيّر، وتنعاد قراءة آخر نسخة بدل ما تنكتب فوق تعديل الأول.
- **التحديث**: الصفحة تقرأ آخر البيانات عند التنقل بين الأقسام (كل دقيقة على الأكثر)، وزر ↻ أعلى الصفحة يحدّث فورًا.
- **مراجع الحالات**: مثل النسخة الأصلية، حقول مرجع الحالة تقبل `DEMO-` فقط عشان ما تنحفظ بيانات مرضى. لو تبين تغيير هذا القرار، يحتاج موافقة الجهة المسؤولة عن خصوصية بيانات المرضى.
- **الإشعارات**: ما فيه إشعارات تلقائية مثل الأصل. ممكن لاحقًا نربط **Power Automate** على قائمة CSSD Orders (مثلًا إيميل أو رسالة Teams عند طلب Rush).

</div>

## For developers

SPFx 1.23 web part, no UI framework.

```bash
cd examples/cssd-sharepoint
npm install          # Node 22.14+ (< 23)
npm run build        # lint + compile + package -> sharepoint/solution/cssd-central-hub.sppkg
npm run start        # local workbench against a real tenant
```

- `src/webparts/cssdCentralHub/app/hubApp.ts` is the UI ported from the original single-file HTML. Inline `on*` handlers were replaced by `data-act` / `data-change` / `data-submit` attributes and one delegated listener, because SharePoint's Content Security Policy blocks inline script.
- `src/webparts/cssdCentralHub/app/store.ts` is the SharePoint REST layer. It creates the lists, loads them, writes only records that changed (with ETag concurrency), and uploads attachments.
- `src/webparts/cssdCentralHub/app/styles.ts` holds the original stylesheet, with every selector scoped under `.cssd-hub` so it does not restyle the SharePoint page.
- After a code change, rebuild and copy `sharepoint/solution/cssd-central-hub.sppkg` to `package/`. Raise the version in `config/package-solution.json` before you upload it again.
