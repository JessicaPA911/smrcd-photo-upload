export interface Tag {
  Id: number;
  Tag: string;
  TagType: string;
}

export const tagFields = {
  LocationTag: 'Location',
  UseTag: 'Use',
  SubjectTag: 'Subject'
} as const;

export type TagField = keyof typeof tagFields | 'ItemTags';
export type TagIds = Record<TagField, number[]>;

export function tagIds(tags: readonly Tag[]): TagIds {
  const ids = (category?: string): number[] =>
    [...new Set(tags.filter(tag => !category || tag.TagType === category).map(tag => tag.Id))];
  return {
    LocationTag: ids('Location'),
    UseTag: ids('Use'),
    SubjectTag: ids('Subject'),
    ItemTags: ids()
  };
}

export function mergeTagIds(existing: TagIds, added: TagIds): TagIds {
  const merge = (field: TagField): number[] => [...new Set([...existing[field], ...added[field]])];
  return {
    LocationTag: merge('LocationTag'),
    UseTag: merge('UseTag'),
    SubjectTag: merge('SubjectTag'),
    ItemTags: merge('ItemTags')
  };
}
