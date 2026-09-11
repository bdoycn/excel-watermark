/**
 * 轻量 OOXML 处理工具。
 *
 * 这里刻意不引入 XML 解析库：xlsx 内部都是很规整的 XML，
 * 用「标签扫描 + 顶层子元素定位」的方式既能精确定位插入点，
 * 又不会在重新序列化时破坏原有内容（格式化、命名空间声明等）。
 */

export const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

export const NS_RELATIONSHIPS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const NS_PACKAGE_RELATIONSHIPS =
  'http://schemas.openxmlformats.org/package/2006/relationships';

/** 转义 XML 文本/属性中的特殊字符 */
export function escapeXml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/** 转义文本节点（属性值以外的场景） */
export function escapeXmlText(value) {
  return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** 反转义（用于读取已有文本后拼装） */
export function unescapeXml(value) {
  return String(value)
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

const TAG_RE = /<(\/)?([A-Za-z_][\w.\-]*:)?([A-Za-z_][\w.\-]*)((?:"[^"]*"|'[^']*'|[^>"'])*?)(\/)?>/g;

/**
 * 扫描 XML，返回根元素及其「直接子元素」的位置信息。
 * 注释 / CDATA / 处理指令会被自然跳过。
 *
 * 返回 { name, prefix, start, openEnd, closeStart, end, openTag, inner, selfClosing, children }
 * children 里每一项结构相同，可继续向下遍历。
 */
export function parseDocument(xml) {
  const stack = [];
  let root = null;
  TAG_RE.lastIndex = 0;
  let match;
  while ((match = TAG_RE.exec(xml)) !== null) {
    const [full, closing, prefix, local, , selfClose] = match;
    const selfClosing = Boolean(selfClose);
    if (closing) {
      const el = stack.pop();
      if (!el) continue;
      el.closeStart = match.index;
      el.end = match.index + full.length;
      el.inner = xml.slice(el.openEnd, el.closeStart);
      if (stack.length === 0) root = el;
    } else {
      const el = {
        name: local,
        prefix: prefix || '',
        start: match.index,
        openEnd: match.index + full.length,
        openTag: full,
        selfClosing,
        children: [],
        closeStart: -1,
        end: -1,
        inner: '',
      };
      if (stack.length > 0) stack[stack.length - 1].children.push(el);
      if (selfClosing) {
        el.closeStart = el.openEnd;
        el.end = el.openEnd;
        if (stack.length === 0) root = el;
      } else {
        stack.push(el);
      }
    }
  }
  if (!root) throw new Error('无法解析 XML：未找到根元素');
  return root;
}

/** 从开标签里取属性值（支持单/双引号），不存在返回 null */
export function getAttr(openTag, name) {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const re = new RegExp(`(?:^|\\s)${escaped}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`);
  const m = re.exec(openTag);
  if (!m) return null;
  return m[1] !== undefined ? m[1] : m[2];
}

/**
 * 取任意命名空间前缀的 id 属性（如 r:id）。
 * 由于正则要求属性名紧跟在空白之后，sheetId（本地名不是 id）不会被误匹配。
 */
export function getRelationshipId(openTag) {
  const re = /\s(?:[A-Za-z_][\w.\-]*:)?id\s*=\s*(?:"([^"]*)"|'([^']*)')/;
  const m = re.exec(openTag);
  if (!m) return null;
  return m[1] !== undefined ? m[1] : m[2];
}

/** 返回开标签里 tag 名之后的属性串（不含结尾的 > 或 />） */
export function tagAttributes(openTag) {
  const m = /^<[A-Za-z_][\w.\-]*(?::[A-Za-z_][\w.\-]*)?((?:"[^"]*"|'[^']*'|[^>"'])*?)\s*\/?>$/.exec(
    openTag,
  );
  return m ? m[1] : '';
}

export function findChild(element, name) {
  return element.children.find((child) => child.name === name) || null;
}

export function findChildren(element, name) {
  return element.children.filter((child) => child.name === name);
}

/**
 * 在根元素的直接子元素中插入片段：
 * 若 order 中的元素已存在，则插入到其中最靠前的一个之前；
 * 否则插入到根元素结束标签之前（这样永远满足 CT_Worksheet 的元素顺序约束）。
 */
export function insertChildOrdered(xml, root, order, snippet) {
  const wanted = new Set(order);
  for (const child of root.children) {
    if (wanted.has(child.name)) {
      return xml.slice(0, child.start) + snippet + xml.slice(child.start);
    }
  }
  if (root.selfClosing) {
    throw new Error(`根元素 <${root.name}/> 为空，无法插入内容`);
  }
  return xml.slice(0, root.closeStart) + snippet + xml.slice(root.closeStart);
}

/** 替换某个子元素（含其内容） */
export function replaceChild(xml, element, snippet) {
  return xml.slice(0, element.start) + snippet + xml.slice(element.end);
}

/** 确保根元素声明了 relationships 命名空间（r: 前缀） */
export function ensureRelationshipsNs(xml) {
  const root = parseDocument(xml);
  if (root.openTag.includes('xmlns:r=')) return xml;
  const openTag = xml.slice(root.start, root.openEnd);
  const patched = /\/>$/.test(openTag)
    ? openTag.replace(/\/>$/, ` xmlns:r="${NS_RELATIONSHIPS}"/>`)
    : openTag.replace(/>$/, ` xmlns:r="${NS_RELATIONSHIPS}">`);
  return xml.slice(0, root.start) + patched + xml.slice(root.openEnd);
}
