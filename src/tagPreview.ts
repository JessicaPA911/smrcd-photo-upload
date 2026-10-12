import { SharePointService, type PreviewTag } from './sharePoint';

const categoryOrder = ['Location', 'Program Area', 'People', 'Practices', 'Subject', 'Use'];

function element<K extends keyof HTMLElementTagNameMap>(
  name: K, props: Partial<HTMLElementTagNameMap[K]> = {}, children: (Node | string)[] = []
): HTMLElementTagNameMap[K] {
  const node = Object.assign(document.createElement(name), props);
  node.append(...children);
  return node;
}

export async function showTagPreview(root: HTMLElement, service: SharePointService, userName: string): Promise<void> {
  let tags: PreviewTag[] = [];
  const selected = new Set<number>();
  let query = '';
  let status = 'Loading your SharePoint tag choices...';
  let loadError = '';

  document.title = 'Photo tags preview';
  root.classList.add('tagPreview');
  const header = element('header', { className: 'previewHeader' }, [
    element('p', { className: 'eyebrow', textContent: 'PHOTO LIBRARY / DESIGN PREVIEW' }),
    element('h1', { textContent: 'Add tags to your photos' }),
    element('p', { textContent: 'Choose tags by category. Combine as many as you need.' }),
    element('p', { className: 'user', textContent: `Signed in as ${userName}` })
  ]);
  const notice = element('div', {
    className: 'previewNotice', textContent:
      'Preview only: selections stay on this page. No photos are uploaded, no tags are saved, and the canvas app is not opened.'
  });
  const search = element('input', {
    type: 'search', placeholder: 'Search categories or tag names', className: 'tagSearch'
  });
  search.id = 'previewTagSearch';
  const searchLabel = element('label', { htmlFor: search.id, textContent: 'Find a tag' });
  const summary = element('section', { className: 'selectionSummary' });
  summary.setAttribute('aria-label', 'Selected tags');
  const grid = element('div', { className: 'categoryGrid' });
  const live = element('p', { className: 'summary', role: 'status' });
  const errorPanel = element('div', { className: 'error', role: 'alert', hidden: true });
  const retry = element('button', {
    type: 'button', className: 'secondary', textContent: 'Retry loading tags', hidden: true
  });
  const content = element('section', { className: 'previewContent' }, [searchLabel, search, live, errorPanel, retry, grid]);
  root.replaceChildren(header, notice, element('div', { className: 'previewLayout' }, [content, summary]));

  function updateSummary(): void {
    summary.replaceChildren(
      element('h2', { textContent: `Selected tags (${selected.size})` }),
      element('p', { textContent: 'Category and tag name stay together, so similar names are easy to distinguish.' })
    );
    const chips = element('ul', { className: 'selectedTagList' });
    tags.filter(tag => selected.has(tag.Id)).forEach(tag => {
      const remove = element('button', {
        type: 'button', className: 'selectedTag',
        onclick: () => { selected.delete(tag.Id); renderChoices(); updateSummary(); }
      }, [element('small', { textContent: tag.TagType }), element('span', { textContent: tag.Tag }), ' Remove']);
      remove.setAttribute('aria-label', `Remove ${tag.TagType}: ${tag.Tag}`);
      chips.append(element('li', {}, [remove]));
    });
    if (!selected.size) { summary.append(element('p', { className: 'emptySelection', textContent: 'Your selected tags will appear here.' })); }
    summary.append(chips, element('button', {
      type: 'button', className: 'secondary', textContent: 'Clear selections', disabled: selected.size === 0,
      onclick: () => { selected.clear(); renderChoices(); updateSummary(); }
    }), element('p', { className: 'user', textContent: 'These selections are a visual preview, not saved metadata.' }));
  }

  function renderChoices(): void {
    grid.replaceChildren();
    const filtered = tags.filter(tag =>
      `${tag.TagType} ${tag.Tag}`.toLocaleLowerCase().includes(query.toLocaleLowerCase().trim()));
    const categories = [...new Set(filtered.map(tag => tag.TagType))].sort((a, b) => {
      const rank = (value: string) => categoryOrder.indexOf(value) < 0 ? categoryOrder.length : categoryOrder.indexOf(value);
      return rank(a) - rank(b) || a.localeCompare(b);
    });
    live.textContent = status || `${filtered.length} active tags in ${categories.length} categories. ${selected.size} selected.`;
    errorPanel.hidden = !loadError;
    errorPanel.textContent = loadError;
    retry.hidden = !loadError;
    categories.forEach(category => {
      const choices = filtered.filter(tag => tag.TagType === category);
      const group = element('fieldset', { className: 'categoryCard' }, [
        element('legend', {}, [category, element('span', { className: 'categoryCount', textContent: String(choices.length) })])
      ]);
      choices.forEach(tag => {
        const check = element('input', { type: 'checkbox', checked: selected.has(tag.Id) });
        const label = element('label', { className: 'tagChoice' }, [
          check, element('span', { className: 'tagChoiceText' }, [
            element('span', { textContent: tag.Tag }),
            ...(tag.description ? [element('small', { textContent: tag.description })] : [])
          ])
        ]);
        check.setAttribute('aria-label', `${category}: ${tag.Tag}`);
        check.onchange = () => {
          if (check.checked) { selected.add(tag.Id); } else { selected.delete(tag.Id); }
          label.classList.toggle('isSelected', check.checked);
          live.textContent = `${filtered.length} active tags in ${categories.length} categories. ${selected.size} selected.`;
          updateSummary();
        };
        label.classList.toggle('isSelected', check.checked);
        group.append(label);
      });
      grid.append(group);
    });
    if (!filtered.length && !status && !loadError) {
      grid.append(element('p', { textContent: tags.length ? 'No matching tags. Try another search.' : 'No active tags are available in this list.' }));
    }
  }

  async function load(): Promise<void> {
    status = 'Loading your SharePoint tag choices...'; loadError = ''; retry.disabled = true; renderChoices();
    try { tags = await service.loadPreviewTags(); }
    catch (e) { loadError = e instanceof Error ? e.message : String(e); console.error(e); }
    finally { status = ''; retry.disabled = false; renderChoices(); updateSummary(); }
  }
  search.oninput = () => { query = search.value; renderChoices(); };
  retry.onclick = () => { void load(); };
  updateSummary();
  await load();
}
