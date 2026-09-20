import * as React from 'react';
import {
  Checkbox,
  ChoiceGroup,
  DefaultButton,
  IChoiceGroupOption,
  Stack,
  Text
} from '@fluentui/react';
import { IDoesFilterPassParams } from 'ag-grid-community';
import { CustomFilterProps, useGridFilter } from 'ag-grid-react';
import styles from './ServiceRequestChoiceFilter.module.scss';

// модель вибраних значень фільтра колонки
interface IRequestChoiceFilterModel {
  values: string[];
}

// параметри фільтра для одиночного або множинного вибору
interface IServiceRequestChoiceFilterProps extends CustomFilterProps<
  unknown,
  unknown,
  IRequestChoiceFilterModel
> {
  allowMultiple?: boolean;
}

const ukrainianCollator = new Intl.Collator('uk-UA');

// збирає унікальні значення поточної колонки
function getFilterOptions(props: CustomFilterProps): string[] {
  const uniqueValues = new Set<string>();

  // додає значення кожного рядка до набору варіантів
  props.api.forEachNode(node => {
    const value = props.getValue(node);
    if (value !== undefined && value !== null && String(value)) {
      uniqueValues.add(String(value));
    }
  });

  const options = Array.from(uniqueValues);
  options.sort(ukrainianCollator.compare);
  return options;
}

// показує множинний вибір значень у меню фільтра колонки
export default function ServiceRequestChoiceFilter(
  props: IServiceRequestChoiceFilterProps
): React.ReactElement {
  const selectedValues = props.model?.values ?? [];
  const options = getFilterOptions(props);

  // перевіряє рядок за вибраними значеннями колонки
  const doesFilterPass = (params: IDoesFilterPassParams): boolean => {
    if (selectedValues.length === 0) {
      return true;
    }

    const value = String(props.getValue(params.node) ?? '');
    return selectedValues.indexOf(value) >= 0;
  };

  useGridFilter({ doesFilterPass });

  const optionControls: React.ReactElement[] = [];
  const singleChoiceOptions: IChoiceGroupOption[] = [];
  for (const option of options) {
    singleChoiceOptions.push({ key: option, text: option });

    // змінює вибір одного значення фільтра
    const handleOptionChange = (
      _event?: React.FormEvent<HTMLElement>,
      checked?: boolean
    ): void => {
      const nextValues: string[] = [];

      for (const selectedValue of selectedValues) {
        if (selectedValue !== option) {
          nextValues.push(selectedValue);
        }
      }

      if (checked) {
        nextValues.push(option);
      }

      props.onModelChange(nextValues.length > 0 ? { values: nextValues } : null);
    };

    optionControls.push(
      <Checkbox
        key={option}
        label={option}
        checked={selectedValues.indexOf(option) >= 0}
        onChange={handleOptionChange}
      />
    );
  }

  // встановлює одне вибране значення фільтра
  const handleSingleChoiceChange = (
    _event?: React.FormEvent<HTMLElement | HTMLInputElement>,
    option?: IChoiceGroupOption
  ): void => {
    props.onModelChange(option ? { values: [option.key] } : null);
  };

  // очищає всі вибрані значення поточної колонки
  const handleClear = (): void => {
    props.onModelChange(null);
  };

  return (
    <div className={styles.root}>
      <Stack tokens={{ childrenGap: 8 }}>
        <Text variant="smallPlus">Оберіть значення</Text>
        <div className={styles.options}>
          {props.allowMultiple ? (
            <Stack tokens={{ childrenGap: 8 }}>{optionControls}</Stack>
          ) : (
            <ChoiceGroup
              options={singleChoiceOptions}
              selectedKey={selectedValues[0]}
              onChange={handleSingleChoiceChange}
            />
          )}
        </div>
        <DefaultButton
          className={styles.clearButton}
          text="Очистити"
          onClick={handleClear}
          disabled={selectedValues.length === 0}
        />
      </Stack>
    </div>
  );
}
