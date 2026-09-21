import * as React from 'react';
import { DefaultButton, Stack, TextField } from '@fluentui/react';
import { IRequestCategory, IRequestSubcategory, IServiceRequestDraft } from '../models/ServiceDeskModels';
import { requestPriorities, requestStatuses } from '../models/ServiceDeskConstants';
import styles from './ServiceRequestGenerator.module.scss';

// вхідні дані тимчасового генератора заявок
interface IServiceRequestGeneratorProps {
  categories: IRequestCategory[];
  subcategories: IRequestSubcategory[];
  disabled: boolean;
  onGenerate: (drafts: IServiceRequestDraft[]) => Promise<void>;
}

// повязана пара категорії та підкатегорії
interface IRequestCategoryPair {
  categoryId: number;
  subcategoryId: number;
}

const generatedUserEmail = 'Oseledko.YY@ua.energy';

const generatedTitles = [
  'Налаштування робочого місця',
  'Проблема з доступом до системи',
  'Перевірка мережевого підключення',
  'Оновлення програмного забезпечення',
  'Діагностика обладнання'
];

const generatedDescriptions = [
  'Потрібна допомога з налаштуванням робочого середовища',
  'Користувач повідомив про проблему під час виконання робочого завдання',
  'Необхідно перевірити доступність сервісу та відновити роботу',
  'Потрібна консультація та технічна перевірка',
  'Необхідно виконати діагностику та запропонувати рішення'
];
const generatedStatuses = [requestStatuses.new, requestStatuses.inProgress];

const generatedPriorities = [
  requestPriorities.low,
  requestPriorities.medium,
  requestPriorities.high
];

// повертає випадковий елемент непорожнього масиву
function getRandomItem<T>(items: T[]): T {
  return items[Math.floor(Math.random() * items.length)];
}

// повертає випадкове ціле число у заданому діапазоні
function getRandomInteger(minimum: number, maximum: number): number {
  return Math.floor(Math.random() * (maximum - minimum + 1)) + minimum;
}

// повертає коректну кількість заявок або позначає помилкове значення
function parseGenerationCount(value: string): number | undefined {
  if (!value.trim()) {
    return 1;
  }

  const count = Number(value);
  const isValidCount = Number.isInteger(count) && count >= 1 && count <= 1000;

  return isValidCount
    ? count
    : undefined;
}

// збирає активні повязані категорії та підкатегорії
function getCategoryPairs(
  categories: IRequestCategory[],
  subcategories: IRequestSubcategory[]
): IRequestCategoryPair[] {
  const pairs: IRequestCategoryPair[] = [];

  for (const category of categories) {
    if (!category.IsActive) {
      continue;
    }

    for (const subcategory of subcategories) {
      const isActiveSubcategoryForCategory =
        subcategory.IsActive && 
        subcategory.CategoryId === category.Id;

      if (isActiveSubcategoryForCategory) {
        pairs.push({ categoryId: category.Id, subcategoryId: subcategory.Id });
      }
    }
  }

  return pairs;
}

// створює випадкові коректні дані тестової заявки
function createRandomDraft(pair: IRequestCategoryPair): IServiceRequestDraft {
  const plannedStart = new Date();
  plannedStart.setHours(plannedStart.getHours() + getRandomInteger(1, 48));
  
  const dueDate = new Date(plannedStart);
  dueDate.setHours(dueDate.getHours() + getRandomInteger(4, 72));
  
  const title = getRandomItem(generatedTitles);

  return {
    title: `${title} ${Date.now()}`,
    description: getRandomItem(generatedDescriptions),
    categoryId: pair.categoryId,
    subcategoryId: pair.subcategoryId,
    status: getRandomItem(generatedStatuses),
    priority: getRandomItem(generatedPriorities),
    requesterIdentity: generatedUserEmail,
    assigneeIdentity: generatedUserEmail,
    plannedStart: plannedStart.toISOString(),
    dueDate: dueDate.toISOString(),
    estimatedHours: getRandomInteger(1, 80) / 2,
    contactEmail: generatedUserEmail,
    requiresOnsiteVisit: Math.random() >= 0.5
  };
}

// показує тимчасову кнопку генерації тестової заявки
export default function ServiceRequestGenerator(
  props: IServiceRequestGeneratorProps
): React.ReactElement {
  const [isGenerating, setIsGenerating] = React.useState(false);
  const [countText, setCountText] = React.useState('');
  const categoryPairs = getCategoryPairs(props.categories, props.subcategories);
  const generationCount = parseGenerationCount(countText);

  // зберігає введену кількість заявок для генерації
  const handleCountChange = (
    _event?: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    value?: string
  ): void => {
    setCountText(value || '');
  };

  // створює задану кількість випадкових заявок після натискання кнопки
  const handleGenerate = async (): Promise<void> => {
    const cannotGenerate = categoryPairs.length === 0 || generationCount === undefined;

    if (cannotGenerate) {
      return;
    }

    const drafts: IServiceRequestDraft[] = [];
    
    for (let index = 0; index < generationCount; index += 1) {
      drafts.push(createRandomDraft(getRandomItem(categoryPairs)));
    }

    setIsGenerating(true);

    try {
      await props.onGenerate(drafts);
    } finally {
      setIsGenerating(false);
    }
  };

  return (
    <Stack horizontal verticalAlign="start" tokens={{ childrenGap: 8 }}>
      <TextField
        type="number"
        min={1}
        max={1000}
        step={1}
        value={countText}
        placeholder="Кількість"
        ariaLabel="Кількість заявок"
        title="Кількість заявок від 1 до 1000"
        className={styles.countField}
        errorMessage={generationCount === undefined ? 'Від 1 до 1000' : undefined}
        disabled={props.disabled || isGenerating}
        onChange={handleCountChange}
      />

      <DefaultButton
        iconProps={{ iconName: 'AddTo' }}
        ariaLabel={isGenerating ? 'Генеруємо заявки' : 'Згенерувати заявки'}
        title={isGenerating ? 'Генеруємо заявки' : 'Згенерувати заявки'}
        className={styles.generateButton}
        onClick={handleGenerate}
        disabled={props.disabled || isGenerating || categoryPairs.length === 0 || generationCount === undefined}
      />
    </Stack>
  );
}
