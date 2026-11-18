export interface AppearanceOption {
  value: string;
  label: string;
}

export interface AppearanceFieldsProps {
  look: string;
  looks: readonly AppearanceOption[];
  language: string;
  languages: readonly AppearanceOption[];
  onLookChange: (look: string) => void;
  onLanguageChange: (language: string) => void;
  translate: (text: string) => string;
  lookLabel?: string;
  languageLabel?: string;
  lookId?: string;
  languageId?: string;
  className?: string;
}

/** Shared, controlled Theme/Language fields for staged application preference dialogs. */
export function AppearanceFields({
  look, looks, language, languages, onLookChange, onLanguageChange, translate,
  lookLabel = 'Theme:', languageLabel = 'Language:', lookId, languageId, className,
}: AppearanceFieldsProps) {
  return <>
    {looks.length > 0 && <label className={className} htmlFor={lookId}>
      {translate(lookLabel)}{' '}<select id={lookId} aria-label={translate(lookLabel).replace(/:$/, '')} value={look} onChange={(event) => onLookChange(event.currentTarget.value)}>
        {looks.map(({ value, label }) => <option key={value} value={value}>{label}</option>)}
      </select>
    </label>}
    {languages.length > 0 && <label className={className} htmlFor={languageId}>
      {translate(languageLabel)}{' '}<select id={languageId} aria-label={translate(languageLabel).replace(/:$/, '')} value={language} onChange={(event) => onLanguageChange(event.currentTarget.value)}>
        {languages.map(({ value, label }) => <option key={value} value={value}>{label}</option>)}
      </select>
    </label>}
  </>;
}
