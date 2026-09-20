import * as React from 'react';
import { DefaultButton } from '@fluentui/react';
import {
  IRequestCategory,
  IRequestSubcategory,
  IServiceRequestDraft
} from '../models/ServiceDeskModels';

// вхідні дані тимчасового генератора заявок
interface IServiceRequestGeneratorProps {
  categories: IRequestCategory[];
  subcategories: IRequestSubcategory[];
  disabled: boolean;
  onGenerate: (draft: IServiceRequestDraft) => Promise<void>;
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
const generatedStatuses = ['Нова', 'В роботі', 'Вирішена', 'Закрита'];
const generatedPriorities = ['Низький', 'Середній', 'Високий'];

// повертає випадковий елемент непорожнього масиву
function getRandomItem<T>(items: T[]): T {
  return items[Math.floor(Math.random() * items.length)];
}

// повертає випадкове ціле число у заданому діапазоні
function getRandomInteger(minimum: number, maximum: number): number {
  return Math.floor(Math.random() * (maximum - minimum + 1)) + minimum;
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
      if (subcategory.IsActive && subcategory.CategoryId === category.Id) {
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
  const categoryPairs = getCategoryPairs(props.categories, props.subcategories);

  // створює одну випадкову заявку після натискання кнопки
  const handleGenerate = async (): Promise<void> => {
    if (categoryPairs.length === 0) {
      return;
    }

    setIsGenerating(true);
    try {
      await props.onGenerate(createRandomDraft(getRandomItem(categoryPairs)));
    } finally {
      setIsGenerating(false);
    }
  };

  return (
    <DefaultButton
      text={isGenerating ? 'Генеруємо' : 'Згенерувати'}
      iconProps={{ iconName: 'TestBeaker' }}
      onClick={handleGenerate}
      disabled={props.disabled || isGenerating || categoryPairs.length === 0}
    />
  );
}
